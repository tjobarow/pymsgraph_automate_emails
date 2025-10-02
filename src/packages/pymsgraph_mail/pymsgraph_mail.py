#!/.venv-linux/bin/ python
# -*-coding:utf-8 -*-
"""
@File    :   SimpleSendMailMsGraph.py
@Time    :   2025/01/27 09:14:11
@Author  :   Thomas Obarowski
@Version :   1.0
@Contact :   tjobarow@gmail.com
@License :   MIT License
@Desc    :   A lightweight wrapper over the MSGraph API that makes it easier
to send simple emails (with option attachments) via Python.
"""
# Modules for SimpleFileAttachment
import base64
import json
import logging
import mimetypes

# Modules for SimpleSendMail
import sys
import time

# Modules for EmailImportance
from enum import Enum
from functools import wraps

import requests


class MsGraphRateLimitExceededError(Exception):
    def __init__(self, message: str, retry_after: int = 90):
        super().__init__(message)
        self.retry_after: int = retry_after


class EmailImportance(str, Enum):
    Low = ("low",)
    Normal = ("normal",)
    High = ("high",)


class BodyType(str, Enum):
    Text = "text"
    HTML = "html"


class SimpleFileAttachment:
    # Required class fields
    ATTACHMENT_FILEPATH: str
    ATTACHMENT_FILENAME: str
    CONTENT_TYPE: str
    FILE_BYTES: bytes
    ENCODED_CONTENT: str

    def __init__(
        self,
        filepath: str | None = None,
        filename: str | None = None,
        filebytes: bytes | None = None,
        content_type: str | None = None,
    ):
        # Validate at least one of the filepath or filebytes parameters are not none.
        if filepath is None and filebytes is None:
            raise ValueError(
                "SimpleFileAttachment class requires either the filepath OR filebytes to not be None, but both parameters were None!"
            )
        elif filepath is not None and filebytes is not None:
            raise RuntimeError(
                "Both a filepath and filebytes were provided, but only one can be used. Please decide which to use and remove the other option."
            )

        # If the filebytes are provided, you also need to specify the filename and content_type
        if filebytes is not None and filename is None:
            raise ValueError(
                "The value of filename must not be None when filebytes are provided, but filename was None."
            )
        if filebytes is not None and content_type is None:
            raise ValueError(
                "The value of content_type must not be None when filebytes are provided, but content_type was None."
            )

        # IF the filepath was provided
        if filepath is not None:
            self.ATTACHMENT_FILEPATH: str = filepath

            # If no filename provided extract from filepath
            if filename is None:
                if "/" in filepath:
                    path_delimiter: str = "/"
                else:
                    path_delimiter: str = "\\"
                # Split the path by slash
                attach_path_arr: list[str] = filepath.split(path_delimiter)
                self.ATTACHMENT_FILENAME: str = attach_path_arr.pop(len(attach_path_arr) - 1)
            # Else use the provided name
            else:
                self.ATTACHMENT_FILENAME: str = filename
            # IF no content type was provided
            if content_type is None:
                # Try to guess the type
                type_guesses: tuple[str] = mimetypes.guess_type(self.ATTACHMENT_FILEPATH)
                # but if no type could be guessed, raise TypeError
                if type_guesses[0] is None:
                    raise TypeError(
                        f"The content type of provided filepath {self.ATTACHMENT_FILEPATH} could not be guessed. Please provide a valid content_type when initalizing the SimpleFileAttachment class."
                    )
                # Else if it could be guessed, set that as the class CONTENT_TYPE
                else:
                    self.CONTENT_TYPE: str = type_guesses[0]
            # Else use the provided content type
            else:
                self.CONTENT_TYPE: str = content_type
            # Load the file into memory as bytes
            try:
                file = open(self.ATTACHMENT_FILEPATH, "rb")
                self.FILE_BYTES: bytes = file.read()
            except FileNotFoundError:
                raise FileNotFoundError(
                    f"SimpleFileAttachment could not locate the file at the provided path: {self.ATTACHMENT_FILEPATH}"
                )
            # B64 encode then utf-8 decode file so it can be sent in JSON body of request
            self.ENCODED_CONTENT: str = base64.b64encode(self.FILE_BYTES).decode("utf-8")

        elif filebytes is not None:
            self.ATTACHMENT_FILENAME: str = filename
            self.CONTENT_TYPE: str = content_type
            self.FILE_BYTES: bytes = filebytes
            self.ENCODED_CONTENT: str = base64.b64encode(self.FILE_BYTES).decode("utf-8")

    def __iter__(self):
        yield "@odata.type", "#microsoft.graph.fileAttachment"
        yield "name", self.ATTACHMENT_FILENAME
        yield "contentType", self.CONTENT_TYPE
        yield "contentBytes", self.ENCODED_CONTENT

    def __dict__(self) -> dict:
        return {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": self.ATTACHMENT_FILENAME,
            "contentType": self.CONTENT_TYPE,
            "contentBytes": self.ENCODED_CONTENT,
        }

    def __str__(self) -> str:
        bytes_removed: dict = self.__dict__().copy()
        bytes_removed.pop("contentBytes")
        return json.dumps(bytes_removed, indent=2)


class SimpleSendMail:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        source_mail_name: str,
        source_mail_address: str,
        oauth_scopes: list = ["https://graph.microsoft.com/.default"],
        verbose: bool = False,
        log_mail_payloads: bool = False,
        max_retries: int = 5,
    ):
        """Initalizes the SimpleSendMail class.

        Args:
            tenant_id (str): Azure Tenant ID to use

            client_id (str): OAuth client id for your Azure enterprise app

            client_secret (str): OAuth client secret for your Azure enterprise app

            source_mail_name (str): Sender (source) name to use, such as "Joe Schmoe"

            source_mail_address (str): Email address to send emails from, such as joeschmoe@mycompany.com

            oauth_scopes (_type_, optional): A list of OAuth scopes to pass the
            MS Graph API OAuth endpoint, if differing from default.
            Defaults to ["https://graph.microsoft.com/.default"].

            verbose (bool, optional): To enable verbose logging, set to True.
            Sets logger to debug level, and enables debug logging to the
            console using a basic logging configuration, if logger did not
            already have console logging enabled. Defaults to False.

        Raises:
            TypeError: Will raise a type error if a provided parameter is not
            the proper type

            ValueError: Will raise a ValueError is a provided parameter is empty
        """
        # Get a logger object, will inherit from calling code if possible
        self._logger: logging.Logger = logging.getLogger(__name__)
        self._max_retries: int = max_retries
        self._log_mail_payloads: bool = log_mail_payloads

        # If verbose was provided
        if verbose:
            self._logger.info("Verbose flag set to True. Enabling debug logging to console.")
            logging.basicConfig(
                level=logging.DEBUG,
                format="%(asctime)s - %(name)s - %(levelname)s - %(lineno)s - %(funcName)20s - %(message)s",
                stream=sys.stdout,
            )
            self._logger.debug("Console debugging set to enabled by verbose=True parameter.")

        # Raise exceptions for non string fields
        if not isinstance(tenant_id, str):
            self._logger.exception(
                f"tenant_id is not of type <str> but rather type {type(tenant_id)}. Raising exception."
            )
            raise TypeError(f"tenant_id must be of type <str>, not {type(tenant_id)}")
        if not isinstance(client_id, str):
            self._logger.exception(
                f"client_id is not of type <str> but rather type {type(client_id)}. Raising exception."
            )
            raise TypeError(f"client_id must be of type <str>, not {type(client_id)}")
        if not isinstance(client_secret, str):
            self._logger.exception(
                f"client_secret is not of type <str> but rather type {type(client_secret)}. Raising exception."
            )
            raise TypeError(f"client_secret must be of type <str>, not {type(client_secret)}")
        if not isinstance(source_mail_name, str):
            self._logger.exception(
                f"source_mail_name is not of type <str> but rather type {type(source_mail_name)}. Raising exception."
            )
            raise TypeError(f"source_mail_name must be of type <str>, not {type(source_mail_name)}")
        if not isinstance(source_mail_address, str):
            self._logger.exception(
                f"source_mail_address is not of type <str> but rather type {type(source_mail_address)}. Raising exception."
            )
            raise TypeError(f"source_mail_address must be of type <str>, not {type(source_mail_address)}")
        if not isinstance(oauth_scopes, list):
            self._logger.exception(
                f"oauth_scopes is not of type <str> but rather type {type(oauth_scopes)}. Raising exception."
            )
            raise TypeError(f"oauth_scopes must be of type <list>, not {type(tenant_id)}")
        for scope in oauth_scopes:
            if not isinstance(scope, str):
                self._logger.exception(
                    f"Scope at index {oauth_scopes.index(scope)} is of type {type(tenant_id)}, not <str>. Raising exception."
                )
                raise TypeError(
                    f"oauth_scopes must be a list of <str> types. Scope at index {oauth_scopes.index(scope)} is of type {type(tenant_id)}, not <str>"
                )

        self._logger.debug("Successfully validated data type of parameters.")

        # Raise exceptions for blank string fields
        if len(tenant_id) <= 0:
            self._logger.exception("tenant_id is an empty string. Raising exception.")
            raise ValueError("tenant_id must not be an empty string.")
        if (
            len(
                client_id,
            )
            <= 0
        ):
            self._logger.exception("client_id is an empty string. Raising exception.")
            raise ValueError("client_id must not be an empty string.")
        if (
            len(
                client_secret,
            )
            <= 0
        ):
            self._logger.exception("client_secret is an empty string. Raising exception.")
            raise ValueError("client_secret must not be an empty string.")
        if (
            len(
                source_mail_name,
            )
            <= 0
        ):
            self._logger.exception("source_mail_name is an empty string. Raising exception.")
            raise ValueError("source_mail_name must not be an empty string.")
        if (
            len(
                source_mail_address,
            )
            <= 0
        ):
            self._logger.exception("source_email_address is an empty string. Raising exception.")
            raise ValueError("source_mail_address must not be an empty string.")
        if len(oauth_scopes) <= 0:
            self._logger.exception("oauth_scopes is an empty list. Raising exception.")
            raise ValueError("oauth_scopes must not be an empty list")
        for scope in oauth_scopes:
            if len(scope) <= 0:
                self._logger.exception(
                    "Scope at index {oauth_scopes.index(scope)} is an empty string. Raising exception"
                )
                raise ValueError(
                    f"oauth_scopes cannot have empty strings present. Scope at index {oauth_scopes.index(scope)} is an empty string."
                )

        self._logger.debug("Successfully validated parameters are populated properly.")

        # Initalize class fields from parameters
        self._tenant_id: str = tenant_id
        self._client_id: str = client_id
        self.__client_secret: str = client_secret
        self._oauth_scopes: list[str] = oauth_scopes
        self._source_mail_name: str = source_mail_name
        self._source_mail_address: str = source_mail_address

        # Initalize oauth_token_info: dict class field by retrieving OAuth token from MSFT
        self.__oauth_token_info: dict[str] = self.__get_OAuth_token()
        self._logger.info("Finished initalizing SimpleSendMail class.")

        self._logger.debug(f"Set class fields with provided parameters: {self.__str__()}")

    def __str__(self):
        return (
            f"Tenant ID: {self._tenant_id}\nClient ID: {self._client_id}"
            + f"\nClient Secret (Is Defined): {True if self.__client_secret is not None else False}"
            + f"\nOAuth Scopes: {self._oauth_scopes}\nSender Name: {self._source_mail_name}"
            + f"\nSender Email Address: {self._source_mail_address}"
        )

    def __get_OAuth_token(self):
        self._logger.debug("Retrieving OAuth token from Microsoft")
        # Construct oauth url using tenant id
        oauth_url: str = f"https://login.microsoftonline.com/{self._tenant_id}/oauth2/v2.0/token"
        self._logger.debug(f"OAuth Token URL set to {oauth_url}")

        # If more than one scope is present in list, join them w/ space delimiter
        scopes: str = " ".join(self._oauth_scopes) if len(self._oauth_scopes) > 1 else self._oauth_scopes[0]

        # Construct OAuth payload
        oauth_body: dict[str] = {
            "grant_type": "client_credentials",
            "client_id": self._client_id,
            "client_secret": self.__client_secret,
            "scope": scopes,
        }

        try:
            # Send post to get oauth token
            response = requests.post(url=oauth_url, data=oauth_body)
            # Used to raise an exception if status code of response is non-200 (2xx)
            response.raise_for_status()

            self._logger.debug(f"Successfully received OAuth token information from {oauth_url}")

            # Save the response to a temporary dictionary
            temp_response_data: dict[str] = response.json()
            # Create a new attribute in that dictionary that is the timestamp the token will expire
            # based on the current unix epoch timestamp + the returned "expires_in" value.
            temp_response_data.update({"expires_at": int(time.time() + temp_response_data["expires_in"])})

            self._logger.debug(
                f"Set new attribute 'expires_at' in OAuth token information to {temp_response_data['expires_at']} - {time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(temp_response_data['expires_at']))}"
            )
            self._logger.info("Successfully retreived OAuth token from Microsoft.")
            return temp_response_data

        except requests.exceptions.RequestException as e:
            self._logger.exception(
                f"A RequestException was raised while trying to retrieve OAuth token from {oauth_url}"
            )
            self._logger.exception(e)
            raise e

    def check_token_validity(func):
        """Wrapper function to check if OAuth token is, or will expire soon
        (5 second buffer) and refresh it with a new one if so.

        Args:
            func (_type_): The function the wrapper decorates

        Returns:
            func (_type_): The function the wrapper decorates
        """

        @wraps(func)
        def check_token_expiration(self, *args, **kwargs):
            # Check if the current seconds timestamp + 5 seconds buffer is greater than the expires_at time.
            if int(time.time() + 5) >= self.__oauth_token_info["expires_at"]:
                self._logger.warning(
                    f"OAuth token is expiring soon at {time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(self.__oauth_token_info['expires_at']))}, retrieving new token"
                )
                self.__oauth_token_info = self.__get_OAuth_token()
            else:
                self._logger.debug("Token not expiring soon.. continuing")
            return func(self, *args, **kwargs)

        return check_token_expiration

    def retry_request(func):
        @wraps(func)
        def wrapper(self, *args, **kwargs):
            req_attempt: int = 0
            while req_attempt < self._max_retries:
                try:
                    self._logger.debug(f"Retry Counter: {req_attempt}/{self._max_retries}")
                    return func(self, *args, **kwargs)
                except MsGraphRateLimitExceededError as error:
                    self._logger.warning(
                        "MSGraph rate limit was exceeded." + f" Retrying in {error.retry_after} seconds..."
                    )
                    # Enable for testing max retries
                    # time.sleep(1)
                    time.sleep(error.retry_after)
                    req_attempt += 1
            self._logger.warning(
                "Max retries was reached, raising " + "MsGraphRateLimitExceededError to calling function."
            )
            raise MsGraphRateLimitExceededError(
                "While attempting to retry request to MS Graph API, "
                + "the maximum number of retries was met without a successful "
                + f"request being made. (Retry Counter: {req_attempt} - "
                + f"Max Retries: {self._max_retries})"
            )

        return wrapper

    @retry_request
    @check_token_validity
    def send_mail(
        self,
        subject: str,
        recipient_emails: str | list[str],
        body_content: str,
        body_type: BodyType = BodyType.Text,
        importance: EmailImportance = EmailImportance.Low,
        attachments: SimpleFileAttachment | list[SimpleFileAttachment] | None = None,
        saveToSentItems: bool = True,
        cc_recipient_emails: list[str] | None = None,
        bcc_recipient_emails: list[str] | None = None,
    ):
        self._logger.info(f"Sending email_details from {self._source_mail_address} to {recipient_emails}")
        self._logger.debug(f"Sending email_details with the following provided parameters: {locals()}")

        mail_url: str = f"https://graph.microsoft.com/v1.0/users/{self._source_mail_address}/sendMail"
        self._logger.debug(f"Constructed mail url {mail_url}")

        headers: dict[str] = {
            "Authorization": f"{self.__oauth_token_info['token_type']} {self.__oauth_token_info['access_token']}",
            "Content-Type": "application/json",
        }

        # Construct the baseline mail payload
        mail_playload: dict = {
            "message": {
                "subject": subject,
                "body": {"contentType": body_type, "content": body_content},
                "toRecipients": [],
                "internetMessageHeaders": [{"name": "x-SentBySimpleSendMail", "value": "true"}],
                "sender": {
                    "emailAddress": {
                        "address": self._source_mail_address,
                        "name": self._source_mail_name,
                    }
                },
                "importance": importance,
            }
        }
        # Add save to sent items if set to true (default is True)
        if saveToSentItems:
            self._logger.info(f"Message will be saved in sent items for {self._source_mail_address}.")
            mail_playload.update({"saveToSentItems": True})

        # Check if recipients is a list and if so, add all recipients, if not just add oen
        if isinstance(recipient_emails, list):
            self._logger.debug(f"Recipients is a list of emails - add all recipients {recipient_emails}")
            for email in recipient_emails:
                mail_playload["message"]["toRecipients"].append({"emailAddress": {"address": email}})
        else:
            self._logger.debug("Recipient emails is a single email_details - adding single email_details")
            mail_playload["message"]["toRecipients"].append({"emailAddress": {"address": recipient_emails}})

        # The next few conditionals will add any additional recipients, set CC and BCC recipients
        if cc_recipient_emails is not None:
            self._logger.info("Message will be sent to list of CC recipients.")
            self._logger.debug(
                f"cc_recipient_emails list was provided - adding CC recipients to payload: {cc_recipient_emails}"
            )
            mail_playload["message"].update({"ccRecipients": []})

            for cc_email in cc_recipient_emails:
                mail_playload["message"]["ccRecipients"].append({"emailAddress": {"address": cc_email}})
            self._logger.debug(f"Added the following CC recipients to mail payload: {cc_recipient_emails}")

        if bcc_recipient_emails is not None:
            self._logger.info("Message will be sent to list of BCC recipients.")
            self._logger.debug(
                f"bcc_recipient_emails list was provided - adding BCC recipients to payload: {bcc_recipient_emails}"
            )
            mail_playload["message"].update({"bccRecipients": []})

            for bcc_email in bcc_recipient_emails:
                mail_playload["message"]["bccRecipients"].append({"emailAddress": {"address": bcc_email}})
            self._logger.debug(f"Added the following BCC recipients to mail payload: {bcc_recipient_emails}")

        # Adding file attachment(s) if any provided.
        if attachments is not None:
            self._logger.info("Email attachment(s) provided and will be attached to message.")
            mail_playload["message"].update({"hasAttachments": True, "attachments": []})
            if isinstance(attachments, list):
                self._logger.debug("Adding multiple file attachments to mail payload.")
                for attachment in attachments:
                    # Make sure it's of type SimpleFileAttachment before trying to add it
                    if not isinstance(attachment, SimpleFileAttachment):
                        self._logger.exception(
                            f"Attachment at index {attachments.index(attachment)} is of type {type(attachment)} but must be of type SimpleFileAttachment."
                        )
                        raise TypeError(
                            f"Attachment at index {attachments.index(attachment)} is of type {type(attachment)} but must be of type SimpleFileAttachment."
                        )
                    mail_playload["message"]["attachments"].append(dict(attachment))
                    self._logger.debug(f"Added attachment {str(attachment)}")
            else:
                # Make sure attachment is of type SimpleFileAttachment
                if not isinstance(attachments, SimpleFileAttachment):
                    self._logger.exception(
                        f"Attachment is of type {type(attachments)} but must be of type SimpleFileAttachment."
                    )
                    raise TypeError(
                        f"Attachment is of type {type(attachments)} but must be of type SimpleFileAttachment."
                    )
                self._logger.debug(
                    f"A single file attachment was provided to function: {attachments.ATTACHMENT_FILENAME}"
                )
                mail_playload["message"]["attachments"].append(dict(attachments))
                self._logger.debug(f"Added single attachment: {str(attachments)}")

        if self._log_mail_payloads:
            self._logger.debug(f"Prepared mail body: {json.dumps(mail_playload,indent=4)}")

        try:
            self._logger.debug("Trying to send mail via MS Graph API")
            # Sending a post request to MS Graph API
            response = requests.post(url=mail_url, headers=headers, json=mail_playload)
            # Checks the status code of response, raises HTTPError if non-2XX
            response.raise_for_status()
            self._logger.info(f"Successfully sent email_details from {self._source_mail_address} to {recipient_emails}")
        # Catches HTTPError that might be raised by raise_for_status()
        except requests.exceptions.HTTPError as http_err:
            # If rate limit was exceeded
            if response.status_code == 429:
                # Do some warning logging
                self._logger.warning(response.text)
                self._logger.warning(
                    "Rate limit was exceeded when trying to email_details "
                    + f"{str(recipient_emails)}. Raising MsGraphRateLimitExceededError..."
                )
                # Raise an instance of MsGraphRateLimitExceededError
                # Which includes the int value from the Retry-After header
                # Which will be used in the wrapper's time.sleep() call
                raise MsGraphRateLimitExceededError(
                    message=str(http_err),
                    retry_after=int(response.headers["Retry-After"]),
                )
            # If the status code was not 429, but something else, raise it
            else:
                self._logger.exception(http_err)
                raise http_err
        except requests.exceptions.RequestException as e:
            self._logger.debug(response.text)
            self._logger.exception(f"An error occurred while attempting to send an email_details to {recipient_emails}")
            self._logger.exception(e)
            raise e

    @retry_request
    @check_token_validity
    def _get_mail_folder(self, folder_name: str, user_principal_name: str) -> dict:
        # TODO Need Mail.ReadWrite
        self._logger.info(f"Retrieving details for {folder_name} folder for user {user_principal_name}")

        request_url: str = (
            f"https://graph.microsoft.com/v1.0/users/{self._source_mail_address}/mailFolders/{folder_name}"
        )
        self._logger.debug(f"Constructed mail url {request_url}")

        headers: dict[str] = {
            "Authorization": f"{self.__oauth_token_info['token_type']} {self.__oauth_token_info['access_token']}",
            "Content-Type": "application/json",
        }
        self._logger.debug("Defined HTTP request headers")

        self._logger.debug(f"Sending get request to {request_url}")

        try:
            response = requests.get(url=request_url)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as err:
            if response.status_code == 401:
                self._logger.exception(
                    f"MSGraph API returned a status code of 401 - Unauthorized. Please validate your application has proper permissions to a mailFolder request."
                )
                self._logger.exception(err)
                raise (err)
            else:
                self._logger.exception(
                    f"An unexcepted HTTP error occurred while attempting to get details for mailFolder {folder_name}"
                )
                self._logger.exception(err)
                raise (err)
        except requests.exceptions.RequestException as err:
            self._logger.exception(
                f"An unexcepted request error occurred while attempting to get details for mailFolder {folder_name}"
            )
            self._logger.exception(err)
            raise (err)

    @retry_request
    @check_token_validity
    def list_message(
        self,
        folder_name: str,
        user_principal_name: str,
        filter: str | None = None,
        search: str | None = None,
        select: str | None = None,
        page_size: int = 10,
        return_count: bool = True,
        adv_query_header: bool = False,
        next_page_url: str | None = None,
    ) -> list:
        self._logger.info(f"Listing messages from {user_principal_name} {folder_name} folder.")

        # Construct default empty params
        params: dict[str:str] = {}
        self._logger.debug("Defined base params")

        headers: dict[str] = {
            "Authorization": f"{self.__oauth_token_info['token_type']} {self.__oauth_token_info['access_token']}",
            "Content-Type": "application/json",
        }
        self._logger.debug("Defined base HTTP request headers")

        if not next_page_url:

            if filter and search:
                raise Exception("You cannot provide both a filter and search parmater to MSGraph API at the same time.")

            if filter:
                self._logger.debug(f"Filter string {filter} supplied. Adding to request parameters.")
                params.update({"$filter": filter})

            if search:
                self._logger.debug(f"Search string {search} supplied. Adding to request parameters")
                params.update({"$search": search})

            if select:
                self._logger.debug(f"Following fields selected to be returned: {select}. Adding to request parameters")
                params.update({"$select": select})

            if page_size:
                self._logger.debug(f"Setting page size to {page_size}")
                params.update({"$top": page_size})

            if return_count:
                self._logger.debug(f"Response to include resource count: {return_count}")
                params.update({"$count": str(return_count).lower()})

            if adv_query_header or filter:
                self._logger.debug(
                    "Flag set to enable advanced query options. Adding header 'ConsistencyLevel=eventual'"
                )
                headers.update({"ConsistencyLevel": "eventual", "$count": "true"})

            request_url: str = (
                f"https://graph.microsoft.com/v1.0/users/{self._source_mail_address}/mailFolders/{folder_name}/messages"
            )
        else:
            request_url: str = next_page_url

        self._logger.debug(f"Sending get request to {request_url} using parameters {params}")

        try:
            # Empty list for returned messages
            messages: list[dict] = []
            # Make inital request for messages
            response = requests.get(url=request_url, params=params, headers=headers)
            response.raise_for_status()

            self._logger.debug(f"API returned {len(response.json().get('value'))} messages.")

            # Add returned messages to list of messages
            messages.extend(response.json().get("value"))

            # This loop will run while the response body contents link to next page of data
            if "@odata.nextLink" in response.json():
                self._logger.debug("Sending recursive function call for additional page of messages from API")
                # Extract next page URL returned my msft
                next_page_url: str = response.json().get("@odata.nextLink")
                # Make recursive call to fetch additional pages of data
                addt_msgs = self.list_message(
                    folder_name=folder_name,
                    user_principal_name=user_principal_name,
                    next_page_url=next_page_url,
                )
                messages.extend(addt_msgs)
                # Then loop and check if the response has link to next page of data again

            self._logger.debug(f"Returning {len(messages)} messages retrieved from the API")
            return messages

        except requests.exceptions.HTTPError as err:
            if response.status_code == 400:
                self._logger.exception(
                    f"MSGraph API returned a status code of 400 - Bad request, and provided an error message of '{response.json().get('error').get('message')}'"
                )
                raise (err)
            if response.status_code == 401:
                self._logger.exception(
                    f"MSGraph API returned a status code of 401 - Unauthorized. Please validate your application has proper permissions to make a listMessages request."
                )
                raise (err)
            else:
                self._logger.exception(
                    f"An unexcepted HTTP error occurred while attempting to get details for mailFolder {folder_name}"
                )
                # self._logger.exception(err)
                raise (err)
        except requests.exceptions.RequestException as err:
            self._logger.exception(
                f"An unexcepted request error occurred while attempting to get details for mailFolder {folder_name}"
            )
            self._logger.exception(err)
            raise (err)

    @retry_request
    @check_token_validity
    def delete_message(self, user_principal_name: str, message_id: str) -> None:
        self._logger.debug(f"Deleting message id {message_id} for user {user_principal_name}")
        # Define headers for API call
        headers: dict[str] = {
            "Authorization": f"{self.__oauth_token_info['token_type']} {self.__oauth_token_info['access_token']}",
            "Content-Type": "application/json",
        }
        self._logger.debug("Defined base HTTP request headers")
        # Define URL
        request_url: str = f"https://graph.microsoft.com/v1.0/users/{user_principal_name}/messages/{message_id}"
        self._logger.debug(f"Defined request URL: {request_url}")

        try:
            response = requests.delete(url=request_url, headers=headers)
            response.raise_for_status()
            self._logger.info(f"Successfully deleted message id {message_id} from mailbox of {user_principal_name}")
        except requests.exceptions.HTTPError as err:
            if response.status_code == 400:
                self._logger.exception(
                    f"MSGraph API returned a status code of 400 - Bad Request, and provided an error message of '{response.json().get('error').get('message')}'"
                )
                raise (err)
            if response.status_code == 401:
                self._logger.exception(
                    f"MSGraph API returned a status code of 401 - Unauthorized. Please validate your application has proper permissions to make delete request."
                )
                raise (err)
            if response.status_code == 404:
                self._logger.exception(
                    f"MSGraph API returned a status code of 404 - Resource Not Found. Please validate the message id for accuracy."
                )
                raise (err)
            else:
                self._logger.exception(
                    f"An unexcepted HTTP error occurred while attempting to delete message {message_id}"
                )
                # self._logger.exception(err)
                raise (err)
        except requests.exceptions.RequestException as err:
            self._logger.exception(
                f"An unexcepted request error occurred while attempting to delete message {message_id}"
            )
            self._logger.exception(err)
            raise (err)
