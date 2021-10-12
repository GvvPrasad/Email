package com.email.testscript;

import java.io.File;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.email.config.ObjectRespo;
import com.email.config.PropertiesFile;
import com.email.utilities.ExcelUtil;

public class Email extends ObjectRespo{	

	@Test(dataProvider = "receiverMails")
	public static void email(String receivermails){	

		// Create object of Property file
		Properties props = new Properties();
		PropertiesFile.GetProperties();

		// this will set host of server
		props.put("mail.smtp.host", "smtp.gmail.com");
		props.put("mail.smtp.socketFactory.port", "587");
		props.put("mail.smtp.socketFactory.class","javax.net.ssl.SSLSocketFactory");
		props.put("mail.smtp.auth", "true");
		props.put("mail.smtp.starttls.enable", "true");
		props.put("mail.smtp.port", "587");

		// This will handle the complete authentication
		Session session = Session.getDefaultInstance(props,
				new javax.mail.Authenticator() {
			protected PasswordAuthentication getPasswordAuthentication() {
				return new PasswordAuthentication(ObjectRespo.senderMail, ObjectRespo.senderPassword);
			}
		});

		try {
			// Create object of MimeMessage class
			Message message = new MimeMessage(session);
			message.addHeader("Content-type", "text/HTML; charset=UTF-8");
			message.addHeader("format", "flowed");
			message.addHeader("Content-Transfer-Encoding", "8bit");
			message.setFrom(new InternetAddress(ObjectRespo.senderMail));
			message.setRecipients(Message.RecipientType.TO,InternetAddress.parse(receivermails));

			// Add the subject 
			message.setSubject(ObjectRespo.mailSubject);

			// Create object to add Attachment
			MimeBodyPart messagecontent = new MimeBodyPart();
			messagecontent.setContent("Hello, <br>\r\n" + 
					"<br>My name is G v v Prasad, <br>\r\n" + 
					"<br>I came across an open QA Position that is available in your organization, I went through Job Description and I find myself fit for it and would like to apply for it. <br>\r\n" + 
					"<br>I started my career as a Manual Tester, then moved to API Testing and started working as both Manual & API Tester. <br>I also have knowlegde on Automation (Selenium with Java). <br>I always try to improve my skill. <br>\r\n" + 
					"<br>\r\n" + 
					"<b>Professional Summary:</b>\r\n" + 
					"<br>Having 2.10 years of experience in the field of Software Testing. <br> Software Test Engineer with experience in <b>\r\n" + 
					"  <i>Manual, API (Postman)</i>\r\n" + 
					"</b> and Agile Methodology.\r\n" + 
					"<br> Having knowledge of <b>Automation (selenium with Java)</b>. <br>\r\n" + 
					"<br>\r\n" + 
					"<b>Total Experience:</b> 2.10 years <br>\r\n" + 
					"<b>Manual Testing:</b> 2.10 years <br>\r\n" + 
					"<b>API Testing:</b> 1 years <br>\r\n" + 
					"<br>\r\n" + 
					"<b>Current CTC:</b> 3.5 LPA <br><br>\r\n" + 
					"\r\n" + 
					"<b>Prefered Location:</b> Hyderbad<br>\r\n" + 
					"\r\n" + 
					"<br>\r\n" + 
					"<b>Skills:</b> Manual, API (Postman), Automation (selenium with Java), TestNG <br>\r\n" + 
					"<br>Do let me know if you need any other information from my end. <br>Please find attached my latest resume for consideration. <br>\r\n" + 
					"<br> Thanks & Regards <br> G v v Prasad <br> 998950812","text/html");

			// Create another object to add Attachment
			MimeBodyPart attachment1 = new MimeBodyPart();
			DataSource source = new FileDataSource(ObjectRespo.resume);
			attachment1.setDataHandler(new DataHandler(source));
			attachment1.setFileName(new File(ObjectRespo.resume).getName());

			// Create object of MimeMultipart class
			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(messagecontent);
			multipart.addBodyPart(attachment1);

			// Set the content
			message.setContent(multipart);

			// Send the email
			Transport.send(message);

		} catch (MessagingException e) {
			System.out.println(receivermails + "  Mail did not Sent");
			throw new RuntimeException(e);
		}
	}

	//Receiver Mails
	@DataProvider
	public Object[][] receiverMails() throws IOException{
		ExcelUtil.getExcel(ObjectRespo.emailList);
		ExcelUtil.getSheet(0);		
		return ExcelUtil.getData();
	}
}