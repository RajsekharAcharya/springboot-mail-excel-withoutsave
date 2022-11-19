package com.vareli.controller.master;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.Query;
import org.hibernate.SessionFactory;
import org.hibernate.transform.Transformers;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.scheduling.annotation.EnableAsync;
import org.springframework.scheduling.annotation.EnableScheduling;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import com.vareli.model.service.CallBeanForCallRegister;

@Configuration
@EnableAsync
@EnableScheduling
@Component
public class AutoMailClassController {
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	
	  @Autowired 
		private JavaMailSender javaMailSender;

	@Autowired
	private SessionFactory sessionFactory;

	public JavaMailSender getJavaMailSender() {
		JavaMailSenderImpl mailSender = new JavaMailSenderImpl();
		mailSender.setHost("smtp.office365.com");
		mailSender.setPort(587);

		mailSender.setUsername("vareli@vareli.co.in");
		mailSender.setPassword("Vtpl@Kol#2103");

		Properties props = mailSender.getJavaMailProperties();
		props.put("mail.smtp.starttls.enable", "true");
		props.setProperty("mail.smtp.ssl.protocols", "TLSv1.2");
		props.put("mail.smtp.port", "587");
		props.put("mail.smtp.host", "smtp.office365.com");
		props.put("mail.smtp.auth", "true");

		return mailSender;
	}

	@Scheduled(cron = "*/10 * * * * *")
	public void sendAutoEmail() throws MessagingException {
		org.hibernate.Session openSession = sessionFactory.openSession();

		List<CallBeanForCallRegister> data = null;
		String sql = "select a.call_type,a.order_code,a.serial_number,a.call_date,a.registration_number,b.customer_name,datediff(current_date,order_code_date) no_days \r\n" + 
				"from trn_call_register a left join tbl_master_account b on (a.cust_id=b.id)\r\n" + 
				"where a.order_code_date <= current_date and a.call_status = 'O' order by datediff(current_date,order_code_date) desc";
		Query q = openSession.createSQLQuery(sql);
		q.setResultTransformer(Transformers.aliasToBean(CallBeanForCallRegister.class));
		data = q.list();

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Service Call");
		Row row = sheet.createRow(0);
		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(11);
		style.setFont(font);
		createCell(row, 0, "Customer Name", style);
		createCell(row, 1, "Ticket Number", style);
		createCell(row, 2, "Order Code", style);
		createCell(row, 3, "Call Start Date", style);
		createCell(row, 4, "No Of Days", style);
		createCell(row, 5, "Description", style);

		writeDataLines(data);

		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
			workbook.write(bos);
		} catch (IOException e) {
			e.printStackTrace();
		}
		byte[] excelFileAsBytes = bos.toByteArray();

		ByteArrayResource resource = new ByteArrayResource(excelFileAsBytes);

		StringBuilder messageBody = new StringBuilder();
		String htmltop = "<div><P> Dear Sir/Mam,<br> Kindly find all the open calls till date.<br>This is an auto generated report</P></div>";
		String htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
		String htmlTableEnd = "</table>";
		String htmlHeaderRowStart = "<tr style=\"background-color:#6FA1D2; color:#ffffff;\">";
		String htmlHeaderRowEnd = "</tr>";
		String htmlTrStart = "<tr style=\"color:#555555;\">";
		String htmlTrEnd = "</tr>";
		String htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
		String htmlTdEnd = "</td>";
		String htmllast = "<div><P>Thanks & Regards,<br>smssupport@vareli.co.in</P></div>";

		String call_type = null;
		String serial_number = null;
		String call_status = null;
		String order_code_date = null;
		String registration_number = null;
		String no_days = null;
		String call_date=null;
		String customer_name=null;
		String order_code=null;

		// Create HTML table header
		messageBody.append(htmltop);
		messageBody.append(htmlTableStart);
		messageBody.append(htmlHeaderRowStart);
		messageBody.append(htmlTdStart + "Customer Name" + htmlTdEnd);
		messageBody.append(htmlTdStart + "Ticket Number" + htmlTdEnd);
		messageBody.append(htmlTdStart + "Order Code" + htmlTdEnd);
		messageBody.append(htmlTdStart + "Call Start Date" + htmlTdEnd);
		messageBody.append(htmlTdStart + "No Of Days" + htmlTdEnd);
		messageBody.append(htmlTdStart + "Description" + htmlTdEnd);
		messageBody.append(htmlHeaderRowEnd);

		for (CallBeanForCallRegister forc : data) {
			
			try {
				customer_name= forc.getCustomer_name().toString();
			} catch (Exception e) {
				// TODO: handle exception
			}
			
			try {
				serial_number=forc.getSerial_number().toString();
			} catch (Exception e) {
				// TODO: handle exception
			}
			try {
				order_code=forc.getOrder_code().toString();
				
			} catch (Exception e) {
				// TODO: handle exception
			}
			try {
				call_date=forc.getCall_date().toString();
			} catch (Exception e) {
				// TODO: handle exception
			}
		
			try {
				no_days = forc.getNo_days().toString();

			} catch (Exception e) {

			}
			try {
				call_type=forc.getCall_type().toString();
			} catch (Exception e) {
				// TODO: handle exception
			}

			messageBody.append(htmlTrStart);
			messageBody.append(htmlTdStart + customer_name + htmlTdEnd);
			messageBody.append(htmlTdStart + serial_number + htmlTdEnd);
			messageBody.append(htmlTdStart + order_code + htmlTdEnd);
			messageBody.append(htmlTdStart + call_date + htmlTdEnd);
			messageBody.append(htmlTdStart + no_days + htmlTdEnd);
			messageBody.append(htmlTdStart + call_type + htmlTdEnd);
			
			messageBody.append(htmlTrEnd);
		}
		messageBody.append(htmlTableEnd);
		messageBody.append(htmllast);
		
		JavaMailSender mail = getJavaMailSender();
		MimeMessage message = mail.createMimeMessage();

		MimeMessageHelper helper = new MimeMessageHelper(message, true);
		helper.setFrom("vareli@vareli.co.in");
		helper.setTo("subhajitg@vareli.co.in");
		helper.setSubject("Open Call List");
		helper.setText(messageBody.toString(),true);
		helper.addAttachment("Open_Call_List.xlsx", resource);
		mail.send(message);

	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
//		if (value instanceof Integer) {
//			cell.setCellValue((Integer) value);
//		} else if (value instanceof Boolean) {
//			cell.setCellValue((Boolean) value);
//		} else {
//			cell.setCellValue((String) value);
//		}
		try {
			cell.setCellValue(value.toString());
		} catch (Exception e) {
			System.out.println("Not Found");
			cell.setCellValue("NA");
		}
		cell.setCellStyle(style);
	}

	private void writeDataLines(List<CallBeanForCallRegister> call) {
		int rowCount = 1;

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);

		for (CallBeanForCallRegister forc : call) {
			Row row = sheet.createRow(rowCount++);
			int columnCount = 0;
			createCell(row, columnCount++, forc.getCustomer_name(), style);
			createCell(row, columnCount++, forc.getSerial_number(), style);
			createCell(row, columnCount++, forc.getOrder_code(), style);
			createCell(row, columnCount++, forc.getCall_date(), style);
			createCell(row, columnCount++, forc.getNo_days(), style);
			createCell(row, columnCount++, forc.getCall_type(), style);

		}
	}

}
