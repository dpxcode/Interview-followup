# Resume Sent Bot

An automated email system for sending job applications and managing follow-ups with HR contacts.

## Features

- Send initial job application emails with resume attachments
- Automated follow-up emails every 2 days
- Resume re-sending every 7 days
- Track email status and responses
- Excel-based contact management
- Detailed follow-up tracking

## Prerequisites

- Node.js (v14 or higher)
- npm (Node Package Manager)
- Gmail account with App Password enabled

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd resume-sent-bot
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file:
```bash
cp .env.example .env
```

4. Edit the `.env` file with your email credentials:
```
EMAIL_USER=your.email@gmail.com
EMAIL_PASS=your_app_password
```

Note: For Gmail, you need to use an App Password:
1. Go to your Google Account settings
2. Enable 2-Step Verification if not already enabled
3. Go to Security > App passwords
4. Generate a new app password for "Mail"
5. Use the generated 16-character password in your .env file

## Setup

1. Create the initial Excel files:
```bash
npm run create-excel
npm run create-followup
```

2. Add HR contacts to `hr_contacts.xlsx` with the following columns:
   - email
   - name
   - company
   - lastSentDate (will be updated automatically)

## Usage

### Sending Initial Emails
```bash
npm start
```
This will send initial application emails to all contacts in the Excel file who haven't been emailed before.

### Sending Follow-up Emails
```bash
npm start -- --follow-up
```
This will send follow-up emails to contacts based on the following schedule:
- Regular follow-up every 2 days
- Resume re-sending every 7 days

### Sending No-Response Follow-ups
```bash
npm start -- --no-response
```
This will send follow-up emails specifically to contacts who:
- Have received an initial application email
- Haven't responded yet (status is 'Pending')
- Haven't received a follow-up in the last 5 days

The no-response follow-up email:
- Has a different subject line indicating it's a follow-up
- Includes your resume
- Emphasizes your continued interest
- Requests feedback or information about next steps

### Checking Email Status
```bash
npm start -- --check-status
```
This will display the current status of all email communications.

### Updating Response Status
```bash
npm start -- --update-status --email "hr@company.com" --status "Interview Scheduled" --notes "Interview scheduled for next week"
```
This will update the status of a specific contact in the tracking file.

### Initializing Tracking
```bash
npm start -- --init-tracking
```
This will initialize the tracking system for existing contacts.

## File Structure

- `index.js` - Main application logic
- `create_excel.js` - Creates the initial HR contacts Excel file
- `create_followup_excel.js` - Creates the follow-up tracking Excel file
- `hr_contacts.xlsx` - Stores HR contact information
- `followup_tracking.xlsx` - Tracks email status and responses
- `.env` - Contains email credentials (not tracked in git)

## Email Templates

The system uses three types of email templates:

1. Initial Application Email:
   - Subject: "Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad"
   - Includes resume attachment
   - Highlights technical skills and experience

2. Regular Follow-up Email:
   - Subject: "Follow-up: Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad"
   - Resume attachment included every 7 days
   - Polite follow-up message

3. No-Response Follow-up Email:
   - Subject: "Following Up: Application for Full Stack Developer (Node.js, React.js, Python) - Durga Parshad"
   - Always includes resume attachment
   - Specifically addresses the lack of response
   - Requests feedback or information about next steps
   - Emphasizes continued interest in the position

## Contributing

Feel free to submit issues and enhancement requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 