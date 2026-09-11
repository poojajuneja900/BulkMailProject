# Bulk Mail Sender

A Java-based application for sending personalized bulk emails with attachments using an Excel spreadsheet as the data source.

## Features

- **Excel Integration**: Reads recipient names and email addresses directly from an `.xlsx` file using Apache POI.
- **Personalized Emails**: Sends customized emails addressing each recipient by their name.
- **Attachments**: Supports attaching files (like a PDF resume/CV) to the outgoing emails.
- **SMTP Support**: Uses JavaMail API to send emails via SMTP (configured for Gmail by default, but customizable).
- **Error Handling**: Skips invalid rows, logs failures, and continues processing the rest of the list.

## Technologies Used

- **Java 11**
- **Maven** (Dependency Management & Build)
- **JavaMail API** (`javax.mail`): For composing and sending emails.
- **Apache POI**: For reading data from Excel files.

## Project Structure

```
BulkMailProject/
├── pom.xml                 # Maven configuration and dependencies
├── data/
│   ├── Recipients.xlsx     # Target Excel file with recipient details
│   └── PoojaResume.pdf     # Attachment file to be sent
└── src/main/java/com/mailer/
    └── BulkMailSender.java # Main application code
```

## Prerequisites

1. **Java Development Kit (JDK) 11** or higher.
2. **Maven** installed.
3. An email account to send from (e.g., Gmail).
4. **App Password**: If using Gmail, you must generate an "App Password" to authenticate safely. Standard passwords won't work with JavaMail.

## Configuration

Before running the application, update the following configurations in `src/main/java/com/mailer/BulkMailSender.java`:

1. **File Paths**:
   - `EXCEL_FILE_PATH`: Absolute path to your `.xlsx` recipient list.
   - `CV_FILE_PATH`: Absolute path to the attachment you want to send.

2. **Email Credentials**:
   - `SENDER_EMAIL`: Your email address.
   - `SENDER_PASSWORD`: Your App Password.
   *(Note: It is highly recommended to use environment variables for credentials rather than hardcoding them into the source).*

3. **Email Content**:
   - Update `MAIL_SUBJECT` with your desired subject line.
   - Modify the `buildMailBody(String name)` method to customize the HTML email body.

### Excel File Format Expected

The application expects an `.xlsx` file where the first row is a header.
- **Column A (Index 0)**: Name
- **Column B (Index 1)**: Email

| Name   | Email              |
|--------|--------------------|
| Sachin | sachin@example.com |
| Rahul  | rahul@example.com  |

## How to Run

You can run the application directly using Maven:

```bash
mvn compile exec:java
```

Or, you can package it into a runnable JAR file:

```bash
# Build the JAR with dependencies
mvn clean package

# Run the generated JAR
java -jar target/BulkMailProject-1.0.0-jar-with-dependencies.jar
```

## Disclaimer

Please ensure you comply with anti-spam laws and regulations in your region when sending bulk emails.
