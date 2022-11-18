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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

@Component
public class mailsent {

    @Autowired 
	private JavaMailSender javaMailSender;

    @Scheduled(cron = "10 * * * * *")
    private void writeHeaderLine() throws MessagingException {

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
    
}
