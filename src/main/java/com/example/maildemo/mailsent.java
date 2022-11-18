package com.example.maildemo;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

@Component
public class mailsent {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;

    @Autowired 
	private JavaMailSender javaMailSender;

    @Scheduled(cron = "10 * * * * *")
    private void writeHeaderLine() throws MessagingException {
        workbook =new XSSFWorkbook();
        sheet = workbook.createSheet("Users");
        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);
        createCell(row, 0, "User ID", style);      
        createCell(row, 1, "E-mail", style);       
        createCell(row, 2, "Full Name", style);    
        createCell(row, 3, "Roles", style);
        createCell(row, 4, "Enabled", style);

        CellStyle style1 = workbook.createCellStyle();
        XSSFFont font1 = workbook.createFont();
        font1.setFontHeight(14);
        style1.setFont(font);

        for (User user : listUsers) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
             
            createCell(row, columnCount++, user.getId(), style1);
            createCell(row, columnCount++, user.getEmail(), style1);
            createCell(row, columnCount++, user.getFullName(), style1);
            createCell(row, columnCount++, user.getRoles().toString(), style1);
            createCell(row, columnCount++, user.isEnabled(), style1);
             
        }



         




        XSSFWorkbook workbook = new XSSFWorkbook();
       List<String> s = Arrays.asList("person","animal");
       CellStyle style = workbook.createCellStyle();
       style.setWrapText(true);
       int n = 1;
       String source = s.get(0);
       Sheet sheet = workbook.createSheet(source);
       for (String i : s) {
           if(!source.equalsIgnoreCase(i)) {
               sheet = workbook.createSheet(i);
               Row header = sheet.createRow(0);

               CellStyle headerStyle = workbook.createCellStyle();

               XSSFFont font = workbook.createFont();
               font.setFontName("Arial");
               font.setBold(true);
               headerStyle.setFont(font);

               Cell headerCell = header.createCell(0);
               headerCell.setCellValue("Name");
               headerCell.setCellStyle(headerStyle);

               headerCell = header.createCell(1);
               headerCell.setCellValue("Age");
               headerCell.setCellStyle(headerStyle);
               source = i;
           }

           Row row = sheet.createRow(n);
           Cell cell = row.createCell(0);
           cell.setCellValue("Udhav Mohata");
           cell.setCellStyle(style);

           cell = row.createCell(1);
           cell.setCellValue(22);
           cell.setCellStyle(style);
           n++;
       }

         ByteArrayOutputStream bos = new ByteArrayOutputStream();
         try {
            workbook.write(bos);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        byte[] excelFileAsBytes = bos.toByteArray();
        

        MimeMessage message = javaMailSender.createMimeMessage();
       MimeMessageHelper helper = new MimeMessageHelper(message, true);
       helper.setFrom("spidercodie@gmail.com");
       helper.setTo("rajsekhar.acharya@gmail.com");
       helper.setSubject("Test Mail");
       helper.setText("Hello world");
       ByteArrayResource resource = new ByteArrayResource(excelFileAsBytes);
       helper.addAttachment("Invoice.xlsx", resource);
       javaMailSender.send(message);
    }
    
    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void writeDataLines(list user) {
        int rowCount = 1;
 
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);
                 
        for (User user : listUsers) {
            Row row = sheet.createRow(rowCount++);
            int columnCount = 0;
             
            createCell(row, columnCount++, user.getId(), style);
            createCell(row, columnCount++, user.getEmail(), style);
            createCell(row, columnCount++, user.getFullName(), style);
            createCell(row, columnCount++, user.getRoles().toString(), style);
            createCell(row, columnCount++, user.isEnabled(), style);
             
        }
    }
}
