# -*-encoding:utf-8 -*-
# noinspection PySingleQuotedDocstring
"""
@File    :   send_emails.py
@Time    :   2025-09-11 12:24:39
@Author  :   Thomas Obarowski
@Contact :   tjobarow@gmail.com
@User    :   tjobarow
@Version :   1.0
@License :   MIT License
@Desc    :   A basic sequential script (top-down execution) that will read a CSV list of end users,
render a provided HTML template for each (using data from the CSV if specified), and send the email to each. Script
does utilize open-source packages written by Thomas to achieve this, though.

Enter the parameter --help to learn about the required parameters
"""

import argparse
import csv
import json

# Standard library imports
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

from dotenv import load_dotenv

# Third-party imports
from loguru import logger

from src.extensions.jinja.jinja_environment_extended import (
    JinjaFileSystemEnvironmentExtended,
)
from src.packages.simple_o365_send_mail_python.simple_o365_send_mail import (
    BodyType,
    EmailImportance,
    SimpleSendMail,
)

#############################################
# Configure argument parsing
#############################################
# Configure arguement parser for the script. Certain config parameters (subject, template path, etc) are passed as args
parser: argparse.ArgumentParser = argparse.ArgumentParser(
    description="Send a rendered HTML template to CSV list of end users."
)

template_help: str = """
Path to HTML template file. 

To render data into template, place Jinja2 variables within the template, using the format {{my_variable}}. 
(See https://documentation.bloomreach.com/engagement/docs/jinja-syntax#variables)

The variable names MUST match their respective CSV column names. For example, to render an end user's first name into
the HTML template (which is the email sent to them), you can have a CSV column first_name, and in the template, place 
the variable {{first_name}} where you'd like the first name to appear. 

If the name of the two variables do not match, it will not work. For example, if the column is titled "user_first_name"
but within the template you have {{first_name}}, nothing will appear for {{first_name}}, as the variable name does not
match the column name.
"""
parser.add_argument("--template_path", "--t", type=str, help=template_help)

csv_help: str = """
Path to the CSV to read. 

MUST contain the column email_address at a MINIMUM, as the script expects to find the email to send to in this column name.
"""
parser.add_argument("--csv_path", "--c", type=str, help=csv_help)

parser.add_argument("--subject", "--s", type=str, help="Subject of email to send.")
parser.add_argument(
    "--email",
    "--e",
    type=str,
    help="Source Email Address of the sending account.",
)
parser.add_argument(
    "--name",
    "--n",
    type=str,
    help="Source Email Name of the sending account.",
)
parser.add_argument(
    "--dry-run",
    "--d",
    action="store_true",
    help="Dry run mode. Renders HTML, but does not send email.",
)
parser.add_argument(
    "--log_level",
    "--l",
    type=str,
    default="INFO",
    choices=["INFO", "WARNING", "DEBUG", "TRACE"],
    help="Log level to use. Choices (by verbosity): INFO, WARNING, DEBUG, TRACE. Defaults to INFO.",
)
env_file_help: str = """
Name of .env file to load. Defaults to .env.

Please format the .env file like:
# MANDATORY VARIABLES
MSGRAPH_TENANT_ID=microsoft tenant id
MSGRAPH_CLIENT_ID=client id
MSGRAPH_CLIENT_SECRET=secret
# OPTIONAL VARIABLES
EMAILS_MUST_MATCH_REGEX=.+@mycompany.com
"""
parser.add_argument("--env-file-name", "--env", type=str, default=".env", help=env_file_help)
parser.add_argument(
    "--max-concurrent-threads",
    "--mct",
    type=int,
    default=10,
    help="Maximum number of concurrent threads to use for sending emails concurrently. Defaults to 10. WARNING: I RECOMMEND NOT INCREASING THE MAX THREADS!",
)
# Parse the provided argumentsw
args = parser.parse_args()

#############################################
# Configure logging
#############################################
# Configure stdout logger
logger.remove()
logger.add(sys.stdout, level=args.log_level, colorize=True)
logger.debug("Successfully configured stdout logger✅.")

# Check if ./logs directory exists, create it if not. If cannot create it, exit
if not os.path.exists("./logs"):
    logger.warning("No './logs' directory found. Creating one before configuring file logger.⚠️")
    try:
        os.mkdir("./logs")
        logger.debug("Successfully created ./logs directory. Continuing⏳...")
    except Exception as err:
        logger.exception(err)
        logger.error("Could not create './logs' directory. Exiting.")
        raise Exception("Failed to created './logs' directory. Exiting.")

# Configure file logger
logger.add(
    sink="./logs/send_emails_py_{time:YYYY-MM-DD}.log",
    rotation="00:00",
    retention="30 days",
    compression="zip",
    level=args.log_level,
    format="{time} - {file} - {function} - {line} - {level} - {message} - {exception}",
    backtrace=True,
    diagnose=False,
)
logger.debug(f"Successfully configured logging to file {logger._core.handlers.get(2)._name} ✅")
logger.info("Logging handlers initialized ✅")

#############################################
# Perform some early on validation on provided configuration
#############################################
# Load environment variables from .env file, validating the file loaded properly
env_loaded: bool = load_dotenv(args.env_file_name, override=True)
logger.debug(f"Successfully loaded env file {args.env_file_name}? {env_loaded}")
if not env_loaded:
    logger.critical(
        "Failed to load environment from '.env' file. Please make sure a .env file exists in the project root."
    )
    raise Exception(
        "Failed to load environment from '.env' file. Please make sure a .env file exists in the project root."
    )

# Validate an email subject is provided
if args.subject is None or args.subject == "":
    logger.critical("No email subject provided. Please provide a subject and try again.")
    raise Exception(
        "No email subject provided. Please provide a subject using parameter --subject or --s and try again."
    )

# Validate template path exists!
if not os.path.exists(args.template_path):
    logger.critical(f"Template path '{args.template_path}' does not exist. Please check path and try again.")
    raise FileNotFoundError(f"Template path '{args.template_path}' does not exist. Please check path and try again.")

# Validate CSV path exists!
if not os.path.exists(args.csv_path):
    logger.critical(f"CSV path '{args.csv_path}' does not exist. Please check path and try again.")
    raise FileNotFoundError(f"CSV path '{args.csv_path}' does not exist. Please check path and try again.")

#######################################################
# Config email client
#######################################################
# Create the credential map from os env variables
ms_creds_map: dict[str, str | None] = {
    "msgraph_tenant_id": os.getenv("MSGRAPH_TENANT_ID"),
    "msgraph_client_id": os.getenv("MSGRAPH_CLIENT_ID"),
    "msgraph_client_secret": os.getenv("MSGRAPH_CLIENT_SECRET"),
}
# Validate that the credentials a non-empty and not null
for key, value in ms_creds_map.items():
    logger.trace(f"Value for {key} is {value}")
    if value is None or value == "":
        logger.critical(
            f"Failed to load {key} from environment. Please make sure it is set in the .env file and not empty."
        )
        raise Exception(
            f"Failed to load {key} from environment. Please make sure it is set in the .env file and not empty."
        )

logger.debug("Initializing SimpleSendMail class ⏳")
try:
    simple_mail: SimpleSendMail = SimpleSendMail(
        tenant_id=ms_creds_map.get("msgraph_tenant_id"),
        client_id=ms_creds_map.get("msgraph_client_id"),
        client_secret=ms_creds_map.get("msgraph_client_secret"),
        source_mail_name=args.name,
        source_mail_address=args.email,
    )
    logger.info("Initalized email client ✅")
except Exception as err:
    logger.exception(err)
    logger.critical("Failed to initialize email client‼📛. Exiting.")
    raise Exception("Failed to initialize email client‼📛. Exiting.")

#######################################################
# Load Jinja Environment
#######################################################
# Create instance of custom JinjaFileSystemEnvironmentExtended class
try:
    jinja_ext_env: JinjaFileSystemEnvironmentExtended = JinjaFileSystemEnvironmentExtended(
        template_file_path=args.template_path
    )
    logger.info("Loaded Jinja2 Environment Extended ✅")
except Exception as err:
    logger.exception(err)
    logger.critical("Failed to load Jinja2 Environment Extended📛. Exiting.")
    raise Exception("Failed to load Jinja2 Environment Extended📛. See previously logged exception.")

#######################################################
# Load provided CSV into list of dictionaries
#######################################################
csv_data: list[dict[str, str]] = []
emails_must_match_regex: str = os.getenv("EMAILS_MUST_MATCH_REGEX")

logger.debug(f"Attempting to open CSV file at path: {args.csv_path}")
try:
    with open(args.csv_path, mode="r", encoding="utf-8") as csv_file:
        logger.debug(f"Successfully opened CSV file at path: {args.csv_path} ✅")
        csv_dict_reader: csv.DictReader = csv.DictReader(csv_file)
        logger.debug("Successfully created DictReader object from CSV file object.✅")
        for i, row in enumerate(csv_dict_reader):
            if (
                emails_must_match_regex is not None
                and emails_must_match_regex != ""
                and not re.match(pattern=emails_must_match_regex, string=row["email_address"])
            ):
                logger.warning(
                    f"Row #{i}-> email_address {row['email_address']} does not match regex {emails_must_match_regex}. Skipping row!"
                )
                logger.warning(f"Skipping row #{i}: {row}")
                continue
            csv_data.append(row)
            logger.trace(f"Loadied Row #{i}: {row}")
    logger.info(f"Loaded {len(csv_data)} rows from CSV ✅")
except Exception as err:
    logger.exception(err)
    err_str: str = "An unhandled error was raised while reading data from CSV file📛. Exiting..."
    logger.critical(err_str)
    raise Exception(err_str)


#######################################################
# Function that will render template and send email - called conncurrently
#######################################################
# Ensures that any exception raised within this function (which will be running in a separate thread) is propagated
# back to the loguru logger
@logger.catch()
def send_email(row: dict[str, str], is_dry_run: bool = False) -> None | dict:
    """
    Send an email to a specified recipient based on the provided template and data.

    This function attempts to send an email using the given row of data. It first renders
    the email template using the data keys and values provided and logs each step of the
    process. Optionally, it can run in "dry-run" mode, in which it renders the template
    but skips the email-sending process.

    :param row: Dictionary containing the data for rendering the email template and
        email address of the recipient.
    :type row: dict[str, str]
    :param is_dry_run: If True, the function will render the email template but will not
        actually send the email.
    :type is_dry_run: bool
    :return: Returns a dictionary containing failure information if an error occurs,
        otherwise returns None.
    :rtype: None | dict
    """
    if is_dry_run:
        logger.warning("⚠️DRY RUN MODE ENABLED. Will render template but not send the email!⚠️")
    logger.info(f"Attempting to send mail for {row['email_address']}")
    logger.debug(f"Rendering template for {row['email_address']}")

    # This function call will render the template with the row's data. Passing it 'row' gives Jinja2 access to all
    # the data loaded from that row of the CSV
    user_rendered_template: str = jinja_ext_env.template.render(row)

    logger.debug(f"Rendered template for {row['email_address']}")

    logger.debug(f"Attempting to send mail to {row['email_address']}")
    if not is_dry_run:
        try:
            simple_mail.send_mail(
                subject=args.subject,
                recipient_emails=row["email_address"],
                body_content=user_rendered_template,
                body_type=BodyType.HTML,
                importance=EmailImportance.High,
            )
            logger.info(f"Successfully sent mail to {row['email_address']}✅")
        except Exception as err:
            logger.exception(err)
            logger.error(f"Failed to send mail to {row['email_address']}⚠️. Logging error information and continuing.")
            failure_information: dict[str, str] = {
                "mail": row["email_address"],
                "row": json.dumps(row),
                "error": str(err),
                "html_content": user_rendered_template,
            }
            return failure_information
    else:
        logger.info(f"Did NOT send mail to {row['email_address']} because --dry-run was provided.")
    return None


#############################################
# Use ThreadPoolExecutor to batch send N emails at a time
# (where N is --max-concurrent-threads -> default is 10
#############################################
logger.debug(f"Creating ThreadPoolExecutor with max_workers={args.max_concurrent_threads}")
with ThreadPoolExecutor(max_workers=args.max_concurrent_threads) as executor:
    # This will call the send_email function against every row of data, 10 at a time. It creates 10 threads at a time
    # to concurrent send 10 emails at a time. We also have to create a list of boolean values representing if dry
    # run is enabled that is N items long, where N is equal to the number of rows of data. This is passed to the
    # is_dry_run parameter of the send_email function.
    #
    # Finally, the result of each function call is saved into a list, results.
    results = list(executor.map(send_email, csv_data, [args.dry_run] * len(csv_data)))

# The send_email function will return None unless it failed to send the email, which instead it returns a small
# dictionary containing failure details. So, if we do list comprehension and create a new list, where the only
# elements are elements of the results list that are not None, we get a list of dictionaries containing failure details.
failures: list[dict] = [result for result in results if result is not None]

if len(failures) > 0:
    filename: str = f"./email_failures_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
    logger.error(
        f"Failed to send emails to {len(failures)} out of {len(csv_data)} users📛‼️. This data will be exported to {filename}"
    )
    with open(filename, mode="w", encoding="utf-8") as failure_file:
        failure_writer = csv.DictWriter(failure_file, fieldnames=failures[0].keys())
        failure_writer.writeheader()
        for failure in failures:
            failure_writer.writerow(failure)
    logger.info(f"Exported {len(failures)} rows of failure information to {filename}")
else:
    logger.info("All emails sent successfully!✅")
