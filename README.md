# Microsoft O365 Emailer

A Python-based tool for sending personalized HTML emails to recipients from a CSV file using Microsoft Graph API (Office
365).

## Features

- **Template-based emails**: Use Jinja2 templates with CSV data for personalization
- **Concurrent sending**: Send multiple emails simultaneously with configurable thread limits
- **Dry-run mode**: Test templates without actually sending emails
- **Email validation**: Optional regex filtering for recipient email addresses
- **Comprehensive logging**: File and console logging with configurable levels
- **Failure tracking**: Automatic CSV export of failed email attempts

## Requirements

- Python 3.12.3
- Microsoft Graph API credentials (Office 365)
- Virtual environment (managed with `virtualenv`)

## Installation

1. Clone the repository and navigate to the project directory
2. Create and activate a virtual environment:

```shell script
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

3. Install dependencies:

```shell script
pip install -r requirements.txt
```

## Configuration

### Environment File (.env)

Create a `.env` file in the project root with the following format:

```
# MANDATORY VARIABLES
MSGRAPH_TENANT_ID=your-microsoft-tenant-id
MSGRAPH_CLIENT_ID=your-client-id
MSGRAPH_CLIENT_SECRET=your-client-secret

# OPTIONAL VARIABLES
EMAILS_MUST_MATCH_REGEX=.+@yourcompany\.com
```

#### Environment Variables Explained:

_Note: The script will look for a a `.env` file in the current directory, unless specified otherwise with the `--env-file-name` parameter._

- **`MSGRAPH_TENANT_ID`** (Required): Your Microsoft Azure tenant ID
- **`MSGRAPH_CLIENT_ID`** (Required): Your registered application's client ID
- **`MSGRAPH_CLIENT_SECRET`** (Required): Your application's client secret
- **`EMAILS_MUST_MATCH_REGEX`** (Optional): Regex pattern to filter recipient emails. Only emails matching this pattern
  will receive emails.

### CSV File Format

Your CSV file **must** contain an `email_address` column. Additional columns can be used for template variables.

Example CSV structure:

```
email_address,first_name,last_name,department,location
user@company.com,John,Doe,IT,New York
```

Additional columns will be passed to the Jinja2 template for rendering.

### HTML Template

Use Jinja2 syntax to create personalized templates. Variable names must match CSV column headers exactly.

Example template:

```html
<!DOCTYPE html>
<html>
<head>
    <title>Is this your location?</title>
</head>
<body>
<h1>Hello {{first_name}} {{Last}}</h1>
<p>Department: {{department}}</p>
<p>You're listed as living in: {{location}}</p>
</body>
</html>
```

Please review the [Jinja2 documentation](https://jinja.palletsprojects.com/en/stable/templates/) for more details on template syntax.

_Note: This package uses `Jinja2==3.1.6`. At the time of writing, this is the stable version. Functionality and documentation may differ in future versions of Jinja2._

## Usage

### Basic Usage

```shell script
python send_emails.py --template_path html_templates/my_template.html --csv_path csv/recipients.csv --subject "Important Request"
```

### Command Line Parameters

| Parameter                  | Short   | Required | Default | Description                                        |
|----------------------------|---------|----------|---------|----------------------------------------------------|
| `--template_path`          | `--t`   | ✅        | -       | Path to HTML template file                         |
| `--csv_path`               | `--c`   | ✅        | -       | Path to CSV file containing recipients             |
| `--subject`                | `--s`   | ✅        | -       | Email subject line                                 |
| `--email`                  | `--e`   | ✅        | -       | Sender email address                               |
| `--name`                   | `--n`   | ✅        | -       | Sender display name                                |
| `--dry-run`                | `--d`   | ❌        | False   | Test mode - render templates but don't send emails |
| `--log_level`              | `--l`   | ❌        | INFO    | Logging verbosity (INFO, WARNING, DEBUG, TRACE)    |
| `--env-file-name`          | `--env` | ❌        | .env    | Name of environment file to load                   |
| `--max-concurrent-threads` | `--mct` | ❌        | 10      | Maximum concurrent email sending threads           |

### Examples

#### Send emails to a few test recipients:

```shell script
python send_emails.py \
  --template_path html_templates/security_alert.html \
  --csv_path csv/test_users.csv \
  --subject "Monthly Security Update" \
  --email "security@company.com" \
  --name "Security Team"
```

#### Dry run to test template rendering:

```shell script
python send_emails.py \
  --template_path html_templates/newsletter.html \
  --csv_path csv/subscribers.csv \
  --subject "Weekly Newsletter" \
  --dry-run \
  --log_level DEBUG
```

#### Send with custom environment file:

```shell script
python send_emails.py \
  --template_path html_templates/alert.html \
  --csv_path csv/recipients.csv \
  --subject "System Maintenance Notice" \
  --env-file-name .env.production
```

## Output and Logging

- **Console logging**: Real-time progress and status updates
- **File logging**: Detailed logs saved to `./logs/send_emails_py_YYYY-MM-DD.log`
- **Failure tracking**: Failed emails exported to `email_failures_YYYY-MM-DD_HH-MM-SS.csv`

_Note: Will create /logs directory in the current directory if it does not already exist._

## Recommended Directory Structure

```
generic_emailer/
├── csv/                    # CSV files with recipient data
├── html_templates/         # Jinja2 HTML email templates
├── logs/                   # Generated log files
├── src/                    # Source code modules
├── .env                    # Environment configuration
├── send_emails.py          # Main script
└── README.md              # This file
```

## Security Notes

- Never commit `.env` files to version control
- Use application-specific passwords for Office 365
- Consider using Azure Key Vault for production deployments
- The `EMAILS_MUST_MATCH_REGEX` helps prevent accidentally sending to external domains

## Troubleshooting

1. **Authentication errors**: Verify your Microsoft Graph API credentials in `.env`
2. **Template rendering issues**: Ensure CSV column names exactly match template variables
3. **Email delivery failures**: Check the generated failure CSV for specific error details
4. **Permission denied**: Ensure your Azure app registration has the necessary Mail.Send permissions (see README.md at
   src\packages\simple_o365_send_mail_python for more help)