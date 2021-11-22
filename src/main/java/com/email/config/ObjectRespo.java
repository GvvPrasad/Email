package com.email.config;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ObjectRespo{

	//Properties file objects
	public static void main(String[] args) {
		PropertiesFile.GetProperties();
	}

	//Project default Variable and Methods
	//Global Variables
	public static String projectPath = System.getProperty("user.dir");
	
	//Excel Utilies Files
	protected static XSSFWorkbook wbFile;
	protected static XSSFSheet shFile;
	protected static XSSFRow row;
	protected static XSSFCell cell;
	protected static FileInputStream dataFile;

	//Properties file
	protected static String emailProperties = projectPath+"//src//main//resources//App.properties";
	protected static String emailList = projectPath+"//src//test//resources//TestEmailList.xlsx";
	protected static String resume = projectPath+"//src//test//resources//Prasad-QA-2.11 Years.docx";

	//Email Variables
	protected static String senderMail;
	protected static String senderPassword;
	protected static String receiverMail;
	protected static String mailSubject;
	protected static String mailContent;
}
