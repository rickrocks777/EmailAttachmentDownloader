package com.emailAttachmentDownloader.EmailAttachmentDownloader;

import jakarta.mail.*;
import jakarta.mail.internet.*;
import java.io.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Timestamp;

public class EmailAttachmentDownloader {

    private static Set<String> downloadedFiles = new HashSet<>();
    private static XSSFWorkbook workbook;
    private static Sheet sheet;
    private static int rowNumber = 0;
    private static final String LOG_FILE_PATH = "C:\\Users\\Anuplab Mukhopadhyay\\Downloads\\attachments\\Logs\\DownloadedFilesLog1.xlsx";
    private static final String BASE_DOWNLOAD_PATH = "C:\\Users\\Anuplab Mukhopadhyay\\Downloads\\attachments\\";

    public static void main(String[] args) {
        try {
            File file = new File(LOG_FILE_PATH);
            if (file.exists()) {
                FileInputStream fis = new FileInputStream(file);
                workbook = new XSSFWorkbook(fis);
                sheet = workbook.getSheetAt(0);
                rowNumber = sheet.getLastRowNum() + 1;
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Downloaded Files Log");
                createHeaderRow();
            }

            // Email credentials
            String username = "anuplab777@gmail.com"; // enter your creds
            String password = "nfvh wttu cdpl sgwc"; //use app specific password for gmail

            // IMAP server settings
            String host = "imap.gmail.com"; //change the host and port according to the email provider, currently set for gmail
            int port = 993;

            Properties props = new Properties();
            props.put("mail.imap.ssl.enable", "true");
            props.put("mail.imap.host", host);
            props.put("mail.imap.port", port);

            Session session = Session.getInstance(props,
                    new jakarta.mail.Authenticator() {
                        protected PasswordAuthentication getPasswordAuthentication() {
                            return new PasswordAuthentication(username, password);
                        }
                    });

            while (true) {
                Store store = session.getStore("imap");
                store.connect();

                Folder inbox = store.getFolder("INBOX");
                inbox.open(Folder.READ_ONLY);

                Message[] messages = inbox.getMessages();
                for (Message message : messages) {
                    Object content = message.getContent();
                    if (content instanceof Multipart) {
                        Multipart multipart = (Multipart) content;
                        for (int i = 0; i < multipart.getCount(); i++) {
                            BodyPart bodyPart = multipart.getBodyPart(i);
                            if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                                String fileName = bodyPart.getFileName();
                                if (!downloadedFiles.contains(fileName) && !isFileDownloaded(fileName)) {
                                    InputStream is = bodyPart.getInputStream();
                                    OutputStream os = new FileOutputStream(new File("C:\\Users\\Anuplab Mukhopadhyay\\Downloads\\attachments\\" + fileName));
                                    byte[] buffer = new byte[4096];
                                    int bytesRead;
                                    while ((bytesRead = is.read(buffer)) != -1) {
                                        os.write(buffer, 0, bytesRead);
                                    }
                                    os.close();
                                    is.close();
                                    System.out.println("Attachment saved: " + fileName);
                                    // Logging the details in Excel sheet
                                    logToFile(message, "C:\\Users\\Anuplab Mukhopadhyay\\Downloads\\attachments\\" + fileName,fileName);
                                    organizeFilesByDate(message.getReceivedDate(), fileName);
                                }
                            }
                        }
                    }
                }

                inbox.close(false);
                store.close();
                Thread.sleep(30000);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createHeaderRow() {
        Row headerRow = sheet.createRow(rowNumber++);
        headerRow.createCell(0).setCellValue("Timestamp");
        headerRow.createCell(1).setCellValue("Sender Email ID");
        headerRow.createCell(2).setCellValue("Subject");
        headerRow.createCell(3).setCellValue("Path");
    }

    private static void organizeFilesByDate(Date receivedDate, String fileName) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(receivedDate);

        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH) + 1; // Month starts from 0
        int day = cal.get(Calendar.DAY_OF_MONTH);

        int folderSuffix = ((day - 1) / 10) + 1; // Determine the folder suffix based on the day

        String folderName = String.format("%04d-%02d-%02d_to_%04d-%02d-%02d", year, month, (folderSuffix - 1) * 10 + 1, year, month, folderSuffix * 10);

        File directory = new File(BASE_DOWNLOAD_PATH + folderName);
        if (!directory.exists()) {
            directory.mkdirs();
        }

        // Move the downloaded file to the corresponding directory
        moveFileToDirectory(BASE_DOWNLOAD_PATH + fileName, directory.getAbsolutePath() + "\\" + fileName);
    }

    private static void moveFileToDirectory(String sourceFilePath, String destinationFilePath) {
        File sourceFile = new File(sourceFilePath);
        File destinationFile = new File(destinationFilePath);
        if (sourceFile.exists() && !destinationFile.exists()) {
            try {
                sourceFile.renameTo(destinationFile);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private static void logToFile(Message message, String path,String filename) {
        try {
            if (sheet == null) {
                sheet = workbook.createSheet("Downloaded Files Log");
                createHeaderRow();
            }

            Row row = sheet.createRow(rowNumber++);
            row.createCell(0).setCellValue(message.getReceivedDate().toString());
            row.createCell(1).setCellValue(Arrays.toString(message.getFrom()));
            row.createCell(2).setCellValue(message.getSubject());
            row.createCell(3).setCellValue(path);
            row.createCell(4).setCellValue(filename);

            FileOutputStream fileOut = new FileOutputStream(LOG_FILE_PATH);
            workbook.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        try {
            createTableIfNotExists();

            String insertSQL = "INSERT INTO t_email_log (from_email , content, attachment_name,attachment_path,Timestamp) VALUES (?, ?, ?, ?, ?)";

            try (Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/stock_statementapp_gcv", "root", "root");
                 PreparedStatement preparedStatement = connection.prepareStatement(insertSQL)) {
                preparedStatement.setString(1, Arrays.toString(message.getFrom()));
                preparedStatement.setString(2, message.getSubject());
                preparedStatement.setString(3, filename);
                preparedStatement.setString(4, path);
                preparedStatement.setTimestamp(5, new Timestamp(message.getReceivedDate().getTime()));
                preparedStatement.executeUpdate();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getTextFromMessage(Message message) throws Exception {
        if (message.getContent() instanceof Multipart) {
            StringBuilder result = new StringBuilder();
            Multipart multipart = (Multipart) message.getContent();
            for (int i = 0; i < multipart.getCount(); i++) {
                BodyPart bodyPart = multipart.getBodyPart(i);
                if (bodyPart.getContentType().contains("text/plain")) {
                    result.append(bodyPart.getContent());
                }
            }
            return result.toString();
        }
        return "";
    }

    private static boolean isFileDownloaded(String fileName) {
        File file = new File("C:\\Users\\Anuplab Mukhopadhyay\\Downloads\\attachments\\" + fileName);
        return file.exists();
    }

    private static void createTableIfNotExists() throws SQLException {
        String createTableSQL = "CREATE TABLE IF NOT EXISTS t_email_log (" +
                "id INT NOT NULL AUTO_INCREMENT," +
                "from_email VARCHAR(255), " +
                "content VARCHAR(255)," +
                "attachment_name VARCHAR(255)," +
                "attachment_path VARCHAR(255)," +
                "Timestamp TIMESTAMP," +
                "ProcessTimestamp TIMESTAMP," +
                "outputfile_name VARCHAR(255)," +
                "outputfile_path VARCHAR(255)," +
                "created_timestamp TIMESTAMP," +
                "created_by VARCHAR(255)," +
                "PRIMARY KEY (id))";

        try (Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/stock_statementapp_gcv", "root", "root");
             PreparedStatement preparedStatement = connection.prepareStatement(createTableSQL)) {
            preparedStatement.executeUpdate();
        }
    }
}
