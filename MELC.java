package MergedEffortLogger;

import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MELC {

	        public static void main(String[] args) throws Exception{

	                FileInputStream file= new FileInputStream(new File("F:\\Semister-2\\AD\\W12\\Authorized Data for Project-3\\Anuj_SD Faculty Effort Log v2.xlsx"));

	                XSSFWorkbook inputWorkbook=new XSSFWorkbook(file);

	                file.close(); //Close the InputStream

	                FileOutputStream output_file =new FileOutputStream(new File("C:\\Users\\User\\Desktop\\MEL.xlsx"));  //Opens Merged excel 

	                inputWorkbook.write(output_file);

	                inputWorkbook.close();

	                output_file.close();    

	        }

	}
