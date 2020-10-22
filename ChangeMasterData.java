import java.awt.GridLayout;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.ChronoField;

import javax.swing.JComboBox;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ChangeMasterData {
	public void ChangeMasterDataMethod() throws EncryptedDocumentException, InvalidFormatException, IOException {
		String StartDate = null, EndDate=null;
		
		InputStream inputStream= new FileInputStream(EmployeeCheck.EmployeeID+".xls");
		Workbook book1 = WorkbookFactory.create(inputStream);
        org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0); 
        
        InputStream myxls = new FileInputStream("Database.xls");
		Workbook book = new HSSFWorkbook(myxls);
		org.apache.poi.ss.usermodel.Sheet EmployeeData = book.getSheetAt(0);
		org.apache.poi.ss.usermodel.Sheet EmployeeDataChange = book.getSheet("Change Data");
		
		int CellID=EmployeeCheck.CellID;
		
		String EmployeeID=EmployeeCheck.EmployeeID;
		String[] EmployeeDataRow= new String[14];
		for(int i=0;i<=13;i++) {
		EmployeeDataRow[i]= EmployeeData.getRow(CellID).getCell(i).toString();
		}		
		int LastRowNumberChange= EmployeeDataChange.getLastRowNum();
		int trigger=0;
		int CellIDChange=LastRowNumberChange+1;
		for(int i=1;i<=LastRowNumberChange;i++) 
		{
			Cell cell2= EmployeeDataChange.getRow(i).getCell(0); 
            String output=cell2.toString();        
			if(output.contentEquals(EmployeeID)) 
			{
				CellIDChange=i;
				trigger=1;
				break;
			}       
		}
			
		
			String[] items = {"","AB", "BC", "MB", "NB", "NF", "NS", "NT", "ON", "PE", "QC","SK", "YT", "NN"};
	          JComboBox<String> combo = new JComboBox<>(items);
	          String[] SH1 = {"","S", "H"};
	          JComboBox<String> combo1 = new JComboBox<>(SH1);          
	          JTextField field1 = new JTextField("");
	          JTextField field3 = new JTextField(""); 
	          JTextField field4 = new JTextField("");
	          JTextField field5 = new JTextField("");
	          JTextField field6 = new JTextField("");
	          JTextField field7 = new JTextField("");
	          JTextField field8 = new JTextField("");
	          JPanel panel = new JPanel(new GridLayout(0, 1));
	          panel.add(new JLabel("Change Effective Date(YYYY-MM-DD)"));
	          panel.add(field1);
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
	          JLabel error = new JLabel("Error Message");
	         
	          int result = JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);	
			
	          if (result == JOptionPane.OK_OPTION) {
	        	  StartDate = field1.getText().toString();
	        	  LocalDate EndDate_Check = LocalDate.parse(PayResults.getRow(PayResults.getLastRowNum()).getCell(2).toString());
	    			EndDate=EndDate_Check.toString();
	    			
	    		if(StartDate.equals("")) {
	    			StartDate=EmployeeDataRow[1];
	    		}
	    		else {
	        	  
	    			LocalDate StartDate_Check = LocalDate.parse(StartDate);
	    			LocalDate HireDate_Check = LocalDate.parse(EmployeeDataRow[1]);
	    			
	    			String day=DayOfWeek.of(StartDate_Check.get(ChronoField.DAY_OF_WEEK)).toString();
	    			while((day.equals("SATURDAY")||day.equals("SUNDAY"))||(StartDate_Check.isBefore(HireDate_Check)||StartDate_Check.isAfter(EndDate_Check))){
	    				panel.add(error);
	    				panel.add(field7);
	      				field7.setText("Employee cannot be hired on weekend, before hire Date");
	      				result= JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
	        	                  JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
	        				if (result != JOptionPane.OK_OPTION) {
	        					JOptionPane.showMessageDialog(null, "Master Data Change cancelled by user");
	        					System.exit(0);
	        				}
	        				panel.remove(error);
	        				panel.remove(field7);
	        				field7.setText("");
	        				StartDate = field1.getText();
	        				StartDate_Check = LocalDate.parse(StartDate);
	        				day=DayOfWeek.of(StartDate_Check.get(ChronoField.DAY_OF_WEEK)).toString();
	    			} 
	    		}
	          }
	          else {
	        	  JOptionPane.showMessageDialog(null, "Master Data Change cancelled by user");
	        	  System.exit(0);
	          }
	         
	          if(trigger==0) {
					Row row = EmployeeDataChange.createRow(CellIDChange);
					for(int i=0;i<=9;i++) 	{
					row.createCell(i).setCellValue(EmployeeDataRow[i]);}
					Cell cell2Update = EmployeeDataChange.getRow(CellIDChange).getCell(1);
					cell2Update.setCellValue(StartDate);
					Cell cell2Update1 = EmployeeDataChange.getRow(CellIDChange).getCell(2);
					cell2Update1.setCellValue(EndDate);
				}
				else {
					for(int i=0;i<=9;i++) {
	    		Cell cell2Update = EmployeeDataChange.getRow(CellIDChange).getCell(i);
	    		cell2Update.setCellValue(EmployeeDataRow[i]);}
					EmployeeDataChange.getRow(CellIDChange).getCell(1).setCellValue(StartDate);
					EmployeeDataChange.getRow(CellIDChange).getCell(2).setCellValue(EndDate);
				}
	          
	          
	           /* EmployeeDataRow[4] = combo.getSelectedItem().toString();
				EmployeeDataRow[5] = combo1.getSelectedItem().toString();
				EmployeeDataRow[6]= field3.getText();						
				EmployeeDataRow[7]= field4.getText();
				EmployeeDataRow[8] = field5.getText();
				EmployeeDataRow[9] = field6.getText();*/
	          
	          String province_empty = combo.getSelectedItem().toString();
			  String salary_empty = combo1.getSelectedItem().toString();
			  String rate_empty= field3.getText();						
				String nof_empty= field4.getText();
				String overtime_empty = field5.getText();
				String garnishment_empty = field6.getText();
				System.out.println(province_empty);
				System.out.println(garnishment_empty);
				
				if(!"".equals(province_empty)) {
					EmployeeDataRow[4] = combo.getSelectedItem().toString();
				}
				if(!"".equals(salary_empty)) {
					EmployeeDataRow[5] = combo1.getSelectedItem().toString();
					
				}
				if(!"".equals(rate_empty)) {
					EmployeeDataRow[6]= field3.getText();
					
				}
				if(!"".equals(nof_empty)) {
					EmployeeDataRow[7]= field4.getText();
					
				}
				if(!"".equals(overtime_empty)) {
					EmployeeDataRow[8] = field5.getText();
					
				}
				if(!"".equals(garnishment_empty)) {
					EmployeeDataRow[9] = field6.getText();
				}
				
				
				
				EmployeeDataRow[11] = "1";
				for(int i=0;i<=13;i++) {
		    		Cell cell2Update = EmployeeData.getRow(CellID).getCell(i);
		    		cell2Update.setCellValue(EmployeeDataRow[i]);}
				JOptionPane.showMessageDialog(null, "Master Data successfully changed");	
	          
			FileOutputStream outputStream = new FileOutputStream("Database.xls");
            book.write(outputStream);
            book.close();
            outputStream.close();
	}

	
}
