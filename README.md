# TPN Daily Account Export

Independent scheduled TPN Browse export. The application downloads the current day's workbook, filters it to the account codes configured in Render, creates one combined XLSX file, and emails it through the configured SMTP mailbox.

## Deployment

Deploy as a Render **Cron Job** using `render.yaml`. The schedule starts at both possible UTC equivalents of 17:00 Europe/London; the application checks local time and only runs at the correct invocation, covering GMT and BST.

All account codes, email addresses, usernames, passwords and tokens must be stored as Render environment variables. Do not commit them to GitHub.

Required variables:

- `TPN_USERNAME`
- `TPN_PASSWORD`
- `ACCOUNT_CODES` (comma-separated)
- `MAIL_USERNAME`
- `MAIL_PASSWORD`
- `MAIL_FROM`
- `MAIL_TO`

The configured Microsoft 365 mailbox must permit authenticated SMTP.
