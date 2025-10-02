
# pymsgraph-mail

A lightweight (requiring only one external dependency - ```requests``` to run) python wrapper over the [user: sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http) API endpoint, allowing you to easily integrate mail functionality into any Python workflow. Supports including file attachments, CC recipients, and BCC recipients.

_Note: While the wrapper does only require ```requests``` to run, you do need to have ```build``` and ```wheel``` installed to build the package from source._

## Requirements

### Python Dependencies

To build and import the package you need to install:

```
requests==2.32.3
```

or rather

```
pip install -r requirements.txt
```

### Microsoft Graph API Requirements
- Enterprise application with OAuth credentials, and permission to use the [SendMail Graph API endpoint](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http#permissions)
- An existing user account to source the email from

## Quick Start

Please check out the [example_usage.py](https://github.com/tjobarow/simple_o365_send_mail_python/blob/7e4fd986a3011da5eba693505c9b1e9decf335bd/example_usage.py) file to learn how to use the package! It's fairly simple overall.

## Custom Message Header
This package implements a custom header attached to each message (email) it sends to the Microsoft Graph API. This is done for the express purpose of making it easier to identify emails sent via this package in other automated workflows. This was implemented as sending tons of emails with this script might result in your mailbox becoming full. You can then write subsequent automation to remove sent messages with this header. 

The header is set as follows::
```python
{"name": "x-SentBySimpleSendMail", "value": "true"}
```
