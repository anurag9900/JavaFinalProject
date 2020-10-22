import javax.swing.*;

import java.awt.GridLayout;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.Math;
import java.util.*;
import java.time.*;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class NewHire {

	public static String EmployeeIDHire, HireDate,TermDate, DOB, Province, SH,
	SHAR, HoursPerWeek, OverTimeHours, Garnishment, PaymentMethod, MDChange, MDChange_for_PP,NoOfLeaveDays  ;
	public static int CellIDHire;	
	
	public static int PP;
	
	public void NewHireMethod() throws IOException, EncryptedDocumentException, InvalidFormatException {
	       
	      EmployeeIDHire=EmployeeCheck.EmployeeID;
	      CellIDHire=EmployeeCheck.CellID;
	      
	      FileInputStream inputStream = new FileInputStream("Database.xls");
          Workbook workbook = WorkbookFactory.create(inputStream);
          org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);
						  
          String[] items = {"AB", "BC", "MB", "NB", "NF", "NS", "NT", "ON", "PE", "QC","SK", "YT", "NN"};
          JComboBox<String> combo = new JComboBox<>(items);
          String[] SH1 = {"S", "H"};
          JComboBox<String> combo1 = new JComboBox<>(SH1);
          String[] PM = {"C", "D"};
          JComboBox<String> combo2 = new JComboBox<>(PM);          
          JTextField field1 = new JTextField("");
          JTextField field2 = new JTextField("");
          JTextField field3 = new JTextField(""); 
          JTextField field4 = new JTextField("");
          JTextField field5 = new JTextField("");
          JTextField field6 = new JTextField("");
          JTextField field7 = new JTextField("");
          JTextField field8 = new JTextField("");
          JTextField field9 = new JTextField("");
          JPanel panel = new JPanel(new GridLayout(0, 1));
          panel.add(new JLabel("Enter Hire Date(YYYY-MM-DD)"));
          panel.add(field1);
          panel.add(new JLabel("Enter Birth Date(YYYY-MM-DD)"));
          panel.add(field2);
          panel.add(new JLabel("Select the Province"));
          panel.add(combo);
          panel.add(new JLabel("Enter Salaried/Hourly Payment"));
          panel.add(combo1);
          panel.add(new JLabel("Enter Monthly Salary/ Hourly Rate"));
          panel.add(field3);
          panel.add(new JLabel("Enter hours worked per week"));
          panel.add(field4);
          panel.add(new JLabel("Enter total Overtime Hours"));
          panel.add(field5);
          panel.add(new JLabel("Enter total Garnishment amount"));
          panel.add(field6);
          panel.add(new JLabel("Select Payment Method"));
          panel.add(combo2);
          JLabel error = new JLabel("Error Message");
         
          int result = JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
          
          if (result == JOptionPane.OK_OPTION) {
            while(true) {
            	panel.remove(error);
				panel.remove(field9);
				field9.setText("");
				
            	try {
        	  HireDate = field1.getText();
  			LocalDate HireDate_Check = LocalDate.parse(HireDate);
  			String day=DayOfWeek.of(HireDate_Check.get(ChronoField.DAY_OF_WEEK)).toString();
  			while(day.equals("SATURDAY")||day.equals("SUNDAY")) {
  				panel.add(error);
  				panel.add(field7);
  				field7.setText("Employee Cannot be hired on Weekend");
  				int result1= JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
  	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
  				if (result1 != JOptionPane.OK_OPTION) {
  					JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
  					System.exit(0);
  				}
  				panel.remove(error);
  				panel.remove(field7);
  				field7.setText("");
  				 HireDate = field1.getText();
  				HireDate_Check = LocalDate.parse(HireDate);
  				day=DayOfWeek.of(HireDate_Check.get(ChronoField.DAY_OF_WEEK)).toString();
  				}
  				DOB = field2.getText();			
  				LocalDate DOB_Check = LocalDate.parse(DOB);						
  				float Age = (float) ChronoUnit.YEARS.between(DOB_Check, HireDate_Check);
  				while(Age<16) {
  					panel.add(error);
  	  				panel.add(field8);
  	  			field8.setText("Employee age less than 16, Hiring terminated");
  	  			int result2= JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
  	  		DOB = field2.getText();			
				DOB_Check = LocalDate.parse(DOB);						
				Age = (float) ChronoUnit.YEARS.between(DOB_Check, HireDate_Check);
  	  			
				if (result2 != JOptionPane.OK_OPTION) {
					JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
					System.exit(0);
				}
				panel.remove(error);
				panel.remove(field8);
				field8.setText("");
  				}
  				Province = combo.getSelectedItem().toString();
				SH = combo1.getSelectedItem().toString();
				SHAR= field3.getText();						
				HoursPerWeek= field4.getText();
				OverTimeHours = field5.getText();
				Garnishment = field6.getText();
				
				float shar1=Float.parseFloat(SHAR);
				float hoursperweek = Float.parseFloat(HoursPerWeek);
				float overtimehours= Float.parseFloat(OverTimeHours);
				float garnishment = Float.parseFloat(Garnishment);
				
				while(shar1<=0 || hoursperweek <=0 || overtimehours<0 || garnishment<0) {
					panel.add(error);
  	  				panel.add(field8);
  	  			field8.setText("Enter the correct number");
  	  		int result2= JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
  	  		
				  	  	SHAR= field3.getText();						
						HoursPerWeek= field4.getText();
						OverTimeHours = field5.getText();
						Garnishment = field6.getText();
						
						shar1=Float.parseFloat(SHAR);
						hoursperweek = Float.parseFloat(HoursPerWeek);
						overtimehours= Float.parseFloat(OverTimeHours);
						garnishment = Float.parseFloat(Garnishment);
  	  		
  	  		
				if (result2 != JOptionPane.OK_OPTION) {
					JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
					System.exit(0);
				}
				panel.remove(error);
				panel.remove(field8);
				field8.setText("");
				}
				
				PaymentMethod = combo2.getSelectedItem().toString();
				MDChange = "0";
				MDChange_for_PP= "0";	
				TermDate="9999-12-31";
				NoOfLeaveDays="0";
				break;
            	}
            	
            	catch(DateTimeParseException e) {
            		panel.add(error);
      				panel.add(field9);
      				field9.setText("enter correct date format");
      				result = JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
      	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE); 
      				
      				if (result != JOptionPane.OK_OPTION) {
      					JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
    					System.exit(0);
      				}
            		
            	}
          }
          }
          
          
          
          
          else {
        	  JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
				System.exit(0);
          }
          
          String[] SettingValue= {EmployeeIDHire, HireDate, TermDate, DOB, Province,
					SH, SHAR, HoursPerWeek, OverTimeHours, Garnishment,
					PaymentMethod, MDChange, MDChange_for_PP,NoOfLeaveDays};
			
			Row row = sheet.createRow(CellIDHire);
			
			for(int i=0; i<=13;i++ ) {
				row.createCell(i).setCellValue(SettingValue[i]);
			}
          
			inputStream.close();
			
			//Modifying the EmployeeData excel file
			 FileOutputStream outputStream = new FileOutputStream("Database.xls");
	            workbook.write(outputStream);
	            workbook.close();
	            outputStream.close();
	         	          
	         Workbook wb = new HSSFWorkbook(); 
	  		 OutputStream os = new FileOutputStream(EmployeeIDHire+".xls"); 
	  	     org.apache.poi.ss.usermodel.Sheet sheet1 = wb.createSheet(EmployeeIDHire);	
	  	   Row row1 = sheet1.createRow(0);
	  	 String[] SettingValue1= {"PP","Start Date","End Date","Gross"
	  			 ,"Over Time","Fed Tax","CPP","EI","Garnishment","Total Deductions"
	  			 ,"ER Benefits","Net payment","Hourly Rate", "Hours", "OverTimeHours", "Leaves", "RetroAmount"};
	  	     
	  	for(int i=0; i<=16;i++ ) {
			row1.createCell(i).setCellValue(SettingValue1[i]);
		}
	  	 wb.write(os);
	  	os.close();
	  	
	  	   JOptionPane.showMessageDialog(null, "Employee is hired, Master Data is updated");
	  	   
	  	   
	  	 int confirm = JOptionPane.showConfirmDialog(null, "Run Payroll?", null, JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE );
			if(confirm==0) 
			{	
				PayrollRun NewHire=new PayrollRun();
				NewHire.PayrollRunMethod();
			}
			else
			{
				JOptionPane.showMessageDialog(null, "Payroll is not run");
				System.exit(0);
			}
	      
	}
}
