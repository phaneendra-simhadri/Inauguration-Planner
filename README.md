# Prototype 2

A simple Python application.

## Project Structure

- `app.py`: Main application entry point.
- `database/`: Directory containing database-related files.
- `static/css/style.css`: Stylesheets for the application.

## Getting Started

1. Ensure you have Python installed on your system.
2. Run the application:
   ```bash
   python app.py
   ```

## Gmail reminder emails

To enable automatic reminder emails for faculty assigned to events tomorrow, set these environment variables before starting the app:

```bash
set EVENT_REMINDER_ENABLED=true
set EVENT_REMINDER_HOUR=18
set EVENT_REMINDER_MINUTE=0
set GMAIL_SMTP_HOST=smtp.gmail.com
set GMAIL_SMTP_PORT=587
set GMAIL_SMTP_USERNAME=your_email@gmail.com
set GMAIL_SMTP_PASSWORD=your_app_password
set GMAIL_FROM_EMAIL=your_email@gmail.com
set GMAIL_FROM_NAME=Event Scheduler
```

The app will send one reminder email per assigned faculty each day at 6:00 PM for events scheduled for tomorrow.

> For Gmail, use a Google App Password instead of your normal account password if two-step verification is enabled.
