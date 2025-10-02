import csv
import io
import os  # Import OS to load HTML email_details templates into string variables

# From simple_o365_send_mail, import the SimpleSendMail class, then import
# SimpleFileAttachment, BodyType, EmailImportance, as needed.
from simple_o365_send_mail import (
    BodyType,
    EmailImportance,
    SimpleFileAttachment,
    SimpleSendMail,
)

"""
Initalize an instance of SimpleSendMail

IMPORTANT: You must have an enterprise application created in Entra ID, which
has the SendMail Graph API permission:
https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http#permissions

Required Parameters:
- tenant_id: Azure Tenant ID
- client_id & client_secret: OAuth Client ID & Secret
- source_mail_name: Source Email Name of the sending account
- source_mail_address: Source Email Address of the sending account

Optional Parameters:
- verbose: Setting verbose to true will initalize the class to use a logger
    set to send debug to console, using the following basic debug configuration:
    
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s - %(levelname)s - %(lineno)s - %(funcName)20s - %(message)s",
        stream=sys.stdout,
    )
    
    Otherwise, the class just creates a new logger object based on it's __name__, 
    which inherits the logging configuration from the Python root logger:
    
    self._logger: logging.Logger = logging.getLogger(__name__)
    
"""
mail_sender = SimpleSendMail(
    tenant_id=os.getenv("azure_tenant_id"),
    client_id=os.getenv("oauth_client_id"),
    client_secret=os.getenv("oauth_client_secret"),
    source_mail_name="Zac Brown",
    source_mail_address="zacbrown@mclaren.com",
    verbose=True,
)

######## BASIC USAGE ########

"""
Basic Usage: Single Recipient, Text Body

Required Parameters:
- subject: Specify email_details subject line
- recipient_emails: Specify a single recipient email_details address as a string
- body_content: body of the mail as a string
- body_type: String value of either "text" or "html". This example uses the
    built-in class BodyType, and references BodyType.Text to specify "text".
- importance: String value of either low, normal, or high. This example uses the
    built-in class EmailImportance, and references EmailImportance.Normal to
    specify "normal".
"""
# BASIC USAGE EXAMPLE #1 - Single recipient, text body
mail_sender.send_mail(
    subject="Tattoo REGRET!!!",
    recipient_emails="landonorris@mclaren.com",
    body_content="I regret getting your first win tattooed on me!!!",
    body_type=BodyType.Text,  # Use the provided BodyType.Text enum to specify the body type as Text
    importance=EmailImportance.Normal,  # Use the provided EmailImportance enum to specify the importance as Normal
)

######## ADVANCED USAGE ########

"""
Advanced Usage: Multiple Recipients, CC'ed addresses, BCC'ed address, HTML email_details body

Note #1: You do not have to load the HTML template from the file system. Any str variable
    containing HTML can be used. This allows you to use templating engines such as 
    Jinja2 to render HTML content dynamically, and then provide it to the send_mail function.

Note #2: The recipient_emails parameter accepts either list of strings (emails), 
    or a single email_details string. However, the cc_recipient_emails and bcc_recipient_emails
    parameters only accept a list of strings (emails). So, even if you only want to
    CC or BCC a single email_details, please present it as a list containing that single email_details
    string.

Required Parameters:
- subject: Specify email_details subject line
- recipient_emails: Specify a multiple recipient email_details addresses as a 
    list of strings
- cc_recipient_emails: A list of email_details addresses as strings with which to CC
    on the email_details
- bcc_recipient_emails: A list of email_details addresses as string with which to BCC 
    on the email_details
- body_content: body of the mail as a string - in this case, the string 
    contents of the loaded HTML file
- body_type: String value of either "text" or "html". This example uses the
    built-in class BodyType, and references BodyType.HTML to specify "html".
- importance: String value of either low, normal, or high. This example uses 
    the built-in class EmailImportance, and references EmailImportance.High to
    specify "high".
"""
# ADVANCED USAGE EXAMPLE #1 - HTML email_details body, CC recipients, BCC recipient,
# and multiple direct recipients
mail_sender.send_mail(
    # Provide a subject
    subject="The secret strategy for beating Red bull!",
    # List of emails to be direct recipients
    recipient_emails=["oscarpiastri@mclaren.com", "landonorris@mclaren.com"],
    # List of email_details(s) to CC
    cc_recipient_emails=[
        "andreastella@mclaren.com",
        "landosengineer@mclaren.com",
        "oscarsengineer@mclaren.com",
    ],
    # List of email_details(s) to BCC
    bcc_recipient_emails=["boardofdirectors@mclaren.com"],
    # The string contents of an HTML file
    body_content=open("super_duper_secret_strategy_for_winning.html", "r").read(),
    # Use the provided BodyType.HTML enum to specify the body type as HTML
    body_type=BodyType.HTML,
    # Use the provided EmailImportance enum to specify the importance as High
    importance=EmailImportance.High,
)

"""
Advanced Usage: Sending one or more attachments

You can also use the built-in SimpleFileAttachment to attach one or more 
local filesystem file(s), as seen below.

The SimpleFileAttachment class supports the following parameters:
Required:
- filepath: path to file to attach

Optional:
- filename (str): optional name for file, if you want it to have a different 
    name than the file currently does. By default, the class pulls the name of 
    the file from the provided path
- content_type (str): Explicitly set the content type for the attachment 
    (such as text/plain). If not specified mimetype.guess_type is used to 
    try to guess the content_type of the file. If none can be resolved,
    a TypeError is raised.

Sending Single Attachment:
    After creating a SimpleFileAttachment, you can simply add it to the attachments
    parameter in the call to send_mail, as seen in example #2. 

Sending Multiple Attachments:
    If you're wanting to send multiple attachments, create an instance of 
    SimpleFileAttachment for each attachment, and pass them all collectively as a list
    of SimpleFileAttachment objects into the attachments parameter, as seen in
    example #3.
"""
# ADVANCED USAGE EXAMPLE #2 - SINGLE FILE ATTACHMENT
mail_sender.send_mail(
    # Provide a subject
    subject="Another file on that super secret strategy",
    # List of emails to be direct recipients
    recipient_emails=["oscarpiastri@mclaren.com", "landonorris@mclaren.com"],
    # List of email_details(s) to CC
    cc_recipient_emails=[
        "andreastella@mclaren.com",
        "landosengineer@mclaren.com",
        "oscarsengineer@mclaren.com",
    ],
    # Basic text body for email_details (can use HTML as well!)
    body_content="Sorry I forgot to send a file! See attached - Zac",
    # Setting body type to text
    body_type=BodyType.Text,
    # Setting the importance to low
    importance=EmailImportance.Low,
    # Adding a SINGLE SimpleFileAttachment to the message
    attachments=SimpleFileAttachment(filepath="./SUPER_SECRET_DO_NOT_SHARE_STRATEGY.txt"),
)

# ADVANCED USAGE EXAMPLE #3 - MULTIPLE FILE ATTACHMENTS
mail_sender.send_mail(
    # Provide a subject
    subject="More files on that super secret strategy",
    # List of emails to be direct recipients
    recipient_emails=["oscarpiastri@mclaren.com", "landonorris@mclaren.com"],
    # List of email_details(s) to CC
    cc_recipient_emails=[
        "andreastella@mclaren.com",
        "landosengineer@mclaren.com",
        "oscarsengineer@mclaren.com",
    ],
    # Load HTML file as the email_details body
    body_content=open("pretty_formatted_email_html_template.html", "r").read(),
    # Setting body type to HTML
    body_type=BodyType.HTML,
    # Setting the importance to High
    importance=EmailImportance.High,
    # Adding MULTIPLE SimpleFileAttachments to the message
    attachments=[
        SimpleFileAttachment(
            filepath="./SUPER_SECRET_DO_NOT_SHARE_STRATEGY.pdf",
            filename="harmless_not_special_file.pdf",
        ),
        SimpleFileAttachment(
            filepath="./STRATEGY_TO_BEAT_FERRARI.pdf",
            filename="this_isnt_a_secret_plan.pdf",
        ),
    ],
)

# ADVANCED USAGE EXAMPLE #4 - CREATING A FILE ATTACHMENT FROM FILE BYTES
"""
If you have a 'file' to attach that does not reside within the local filesystem,
such as one you created as a bytes object in memory, you can provide the
SimpleFileAttachment class the bytes via the file_bytes parameter to create an
attachment based on those bytes.

If you use the file_bytes parameter, you MUST provide a valid filename and
content_type parameter, and CANNOT provide a filepath parameter.
"""
# This 2d list will be converted to an in-memory CSV object
test_csv_data: list = [
    ["Name", "Role"],
    ["Lando Norris", "F1 Driver"],
    ["Oscar Piastri", "F1 Driver"],
    ["Andrea Stella", "Small Boss Man"],
    ["Zac Brown", "Big Boss Man"],
]
# Creating an empty BytesIO object in memory
csv_output = io.BytesIO()
# Creating a TextIOWrapper and passing it to the csv.writer function
csv_writer = csv.writer(io.TextIOWrapper(csv_output, encoding="utf-8", newline=""))
# Writing the list data to the 'in-memory CSV' object
csv_writer.writerows(test_csv_data)
# Getting the bytes for the 'in-memory csv' object.
# The SimpleFileAttachment class will encode the bytes using utf-8 for you
csv_bytes = csv_output.getvalue()

# Creating a SimpleFileAttachment using the csv_bytes, which requires you to
# provide a filename and content_type
SimpleFileAttachment(filebytes=csv_bytes, filename="mclaren_employees.csv", content_type="text/csv")
