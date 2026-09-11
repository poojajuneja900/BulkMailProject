package com.mailer;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * BulkMailSender
 * ---------------
 * Reads "Name" and "Email" columns from an Excel sheet and sends
 * a personalized email with a CV attachment to each recipient.
 *
 * Excel format expected (first row = header):
 * -------------------------------------------
 * | Name        | Email                     |
 * -------------------------------------------
 * | Sachin      | sachin@example.com        |
 * | Rahul       | rahul@example.com         |
 * -------------------------------------------
 *
 * Required libraries (add to your project / classpath):
 *  - javax.mail (jakarta.mail-1.6.x or javax.mail-1.6.2.jar)
 *  - Apache POI (poi-5.x.jar, poi-ooxml-5.x.jar and their dependencies)
 *
 * Maven dependencies (if using Maven):
 *  <dependency>
 *      <groupId>com.sun.mail</groupId>
 *      <artifactId>javax.mail</artifactId>
 *      <version>1.6.2</version>
 *  </dependency>
 *  <dependency>
 *      <groupId>org.apache.poi</groupId>
 *      <artifactId>poi-ooxml</artifactId>
 *      <version>5.2.5</version>
 *  </dependency>
 */
public class BulkMailSender {

    // ==== CONFIGURE THESE ====
    private static final String EXCEL_FILE_PATH = "C:\\Users\\pooja.j\\Downloads\\BulkMailProjectFull\\BulkMailProject\\data\\Recipients.xlsx"; // Excel with Name & Email
    private static final String CV_FILE_PATH = "C:\\Users\\pooja.j\\Downloads\\BulkMailProjectFull\\BulkMailProject\\data\\PoojaResume.pdf"; // Your CV file

    // NOTE: Recommended to move these two to environment variables instead of
    // hardcoding, e.g. System.getenv("MAIL_USER") / System.getenv("MAIL_PASS"),
    // so the app password never sits in source code or version control.
    private static final String SENDER_EMAIL = "jjnsjspoooja@gmail.com";
    private static final String SENDER_PASSWORD = "vbtlhkmkmrjdevvb"; // Use App Password, not your real password
    private static final String MAIL_SUBJECT = "Application for Software Engineer | Gen AI, RAG & AI Agents";

    // SMTP settings for Gmail (change if using Outlook/other provider)
    private static final String SMTP_HOST = "smtp.gmail.com";
    private static final String SMTP_PORT = "587";

    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Session session = createMailSession();

            int successCount = 0;
            int failCount = 0;

            // Start from row 1 to skip header row (row 0)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String name = getCellValue(row.getCell(0));
                String email = getCellValue(row.getCell(1));

                if (email == null || email.trim().isEmpty()) {
                    System.out.println("Row " + (i + 1) + ": Skipped (no email found)");
                    continue;
                }

                try {
                    sendMail(session, name, email);
                    System.out.println("Row " + (i + 1) + ": Mail sent to " + email + " (" + name + ")");
                    successCount++;
                } catch (Exception e) {
                    System.out.println("Row " + (i + 1) + ": FAILED for " + email + " -> " + e.getMessage());
                    failCount++;
                }
            }

            System.out.println("\n=== Mail Sending Completed ===");
            System.out.println("Success: " + successCount + " | Failed: " + failCount);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Creates and returns a Mail Session configured for SMTP with TLS.
     */
    private static Session createMailSession() {
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", SMTP_HOST);
        props.put("mail.smtp.port", SMTP_PORT);

        return Session.getInstance(props, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(SENDER_EMAIL, SENDER_PASSWORD);
            }
        });
    }

    /**
     * Builds and sends a single email with attachment.
     */
    private static void sendMail(Session session, String name, String toEmail) throws MessagingException, IOException {
        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(SENDER_EMAIL));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(toEmail));
        message.setSubject(MAIL_SUBJECT);

        String mailBody = buildMailBody(name);

        MimeBodyPart textPart = new MimeBodyPart();
        textPart.setContent(mailBody, "text/html; charset=utf-8");

        // ---- CV Attachment ----
        MimeBodyPart attachmentPart = new MimeBodyPart();
        attachmentPart.attachFile(new File(CV_FILE_PATH));

        MimeMultipart multipart = new MimeMultipart();
        multipart.addBodyPart(textPart);
        multipart.addBodyPart(attachmentPart);

        message.setContent(multipart);

        Transport.send(message);
    }

    /**
     * Builds the mail body text (HTML). "name" is the first name picked from
     * the Excel row and used directly in the greeting.
     */
    private static String buildMailBody(String name) {
        String greetingName = (name == null || name.trim().isEmpty()) ? "HR Team" : name;

        return "Hi " + greetingName + ",<br><br>"
        		+ "I hope you are doing well.<br><br>"
                + "I am <b>Pooja Juneja</b>, currently working as a Software Engineer at Alert Enterprise with 2 years of experience "
                + "in software development, test automation, and enterprise systems.<br><br>"
                + "Along with my core experience in Java, Python, JavaScript, Selenium, Playwright, TestNG, Maven, and SQL databases, "
                + "I have actively expanded my skills into Generative AI. I possess hands-on expertise in Model Context Protocol (MCP) server development, "
                + "Guardrails integration, Schema Discovery tooling, RAG architecture, and automated error-recovery systems.<br><br>"
                + "I am looking for Software Engineer, SDET, or GenAI roles where I can combine both my automation and AI engineering capabilities. "
                + "If there are any open roles, I would be grateful if you could consider my profile.<br><br>"
                + "Please find my attached resume for your reference.<br><br>"
                + "Best regards,<br>"
                + "<b>Pooja Juneja</b><br>"
                + "Software Engineer | GenAI & Automation<br>"
                + "📞 +91 88377 78196<br>"
                + "✉️ <a href=\"mailto:jjnsjspoooja@gmail.com\">jjnsjspoooja@gmail.com</a><br>"
                + "LinkedIn: <a href=\"https://www.linkedin.com/in/pooja-juneja-801104232/\">linkedin.com/in/pooja-juneja-801104232</a>";   
    }

    /**
     * Reads a cell value as String regardless of its underlying type (text/number).
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());
            default:
                return cell.toString().trim();
        }
    }
}