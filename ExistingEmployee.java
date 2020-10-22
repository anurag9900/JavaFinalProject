import java.awt.Color;
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

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExistingEmployee {

	public static int ChangeMasterDataTrigger=0;
	public void ExistingEmployeeMethod() throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		InputStream inputStream= new FileInputStream(EmployeeCheck.EmployeeID+".xls");
		//Workbook book1 = new HSSFWorkbook(myxls1);
		//org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0);
		Workbook book1 = WorkbookFactory.create(inputStream);
        org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0); 
        
        InputStream myxls = new FileInputStream("Database.xls");
		Workbook book = new HSSFWorkbook(myxls);
		org.apache.poi.ss.usermodel.Sheet EmployeeData = book.getSheetAt(0);
		
		
		String[] items = {"Display Payresults","Run payroll","Apply Leave","Change MasterData", "Terminate EE"};
        JComboBox<String> combo = new JComboBox<>(items);
        JPanel panel = new JPanel(new GridLayout(0, 1));
        panel.add(new JLabel("Select Employee Action"));
        panel.add(combo);
        
        int result = JOptionPane.showConfirmDialog(null, panel, "Enter the EE details",
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        
        if (result == JOptionPane.OK_OPTION) {
        	try {
        	String ActionValue=combo.getSelectedItem().toString();
        	
        	if(ActionValue.equals("Display Payresults")) 
        		{        
                
                if(PayResults.getLastRowNum()==0) {
                	
                	JOptionPane.showMessageDialog(null, "No Payresults to Display");
                	int confirm = JOptionPane.showConfirmDialog(null, "Run Payroll?", null, JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE );
        			if(confirm==0) 
        			{	
        				PayrollRun ExistingEE=new PayrollRun();
        				ExistingEE.PayrollRunMethod();
        			}
        			else
        			{
        				JOptionPane.showMessageDialog(null, "Payroll is not run");
        				System.exit(0);
        			}
                }
        		
        		
                String[] PayPeriodNumber = new String[PayResults.getLastRowNum()+1];
        		String [] PayPeriodStartDate = new String[PayResults.getLastRowNum()+1];
        		String [] PayPeriodEndDate = new String[PayResults.getLastRowNum()+1];
        		for(int i=1;i<=PayResults.getLastRowNum();i++) {
        			PayPeriodNumber[i]=PayResults.getRow(i).getCell(0).toString();
        			PayPeriodStartDate[i]=PayResults.getRow(i).getCell(1).toString();
        			PayPeriodEndDate[i]=PayResults.getRow(i).getCell(2).toString();
        			}
        			
        			String output="Enter Pay Period of Payresults\n";
        			for(int i=1;i<=PayResults.getLastRowNum();i++) {
        				output=output+ "Period:"+PayPeriodNumber[i]+"   "+
        						"Start Date:"+ PayPeriodStartDate[i]+"   "+
        						"End Date:"+PayPeriodEndDate[i]+"\n";
        						}
        			String PP=(JOptionPane.showInputDialog(output));
        			int pp1=Integer.parseInt(PP);
        			int pp2=Integer.parseInt(PayPeriodNumber[1]);
        			
        			while(pp1<pp2 && pp1>12)
        			{String output1="Incorrect Number, enter correct PP number"+"\n";
        			output1=output1+output;
        			PP=(JOptionPane.showInputDialog(output1));
        			pp1=Integer.parseInt(PP);
        			}
        			
        			
        			DisplayPaySlips DisplayPaySlips1=new DisplayPaySlips();
        			DisplayPaySlips1.GeneratePaySlip(PP);
        			}
        	
        	if(ActionValue.equals("Apply Leave")) {
        		
        		int cell_Num= EmployeeCheck.CellID;
        		String leaves=JOptionPane.showInputDialog("Enter Total Leaves in Days ");
        		if(leaves.equals(null)) {
        			JOptionPane.showMessageDialog(null, "Action Cancelled by user");
        			System.exit(0);
        		}
        		Cell cell2Update = EmployeeData.getRow(cell_Num).getCell(13);
        		cell2Update.setCellValue(leaves);
        		
        		FileOutputStream outputStream = new FileOutputStream("Database.xls");
	            book.write(outputStream);
	            book.close();
	            outputStream.close();
	            
	            JOptionPane.showMessageDialog(null, "Employee Master Data is updated");
        	}
        	
        	if(ActionValue.equals("Run payroll")) {
        		
        		PayrollRun payroll=new PayrollRun();
        		payroll.PayrollRunMethod();
        		
        	}
        	 		
        	if(ActionValue.equals("Terminate EE")) {
        		
        		int cell_Num= EmployeeCheck.CellID;        	
        		
        		LocalDate LastDate = null;
        		
        		LocalDate EEhireDate=LocalDate.parse(EmployeeData.getRow(cell_Num).getCell(1).toString());
        		if(PayResults.getLastRowNum()!=0) {
        		LastDate=LocalDate.parse(PayResults.getRow(PayResults.getLastRowNum()).getCell(2).toString());}
        		else LastDate=EEhireDate;
        		
        		String TermDate=JOptionPane.showInputDialog("Terminate EE after:"+LastDate+"\n"+"Termination Date - YYYY-MM-DD ");
        		LocalDate TerminationDate=LocalDate.parse(TermDate);
        		String day=DayOfWeek.of(TerminationDate.get(ChronoField.DAY_OF_WEEK)).toString();
        		
        		while(TerminationDate.isBefore(EEhireDate)||TerminationDate.isBefore(LastDate)||day.equals("SATURDAY")||day.equals("SUNDAY")) {
        			
        			if(TerminationDate.isBefore(EEhireDate)) {
        				TermDate=JOptionPane.showInputDialog("EE cannot be terminated before HireDate:"+EEhireDate+"\n"
        						+ "Enter Termination Date - YYYY-MM-DD ");
        				TerminationDate=LocalDate.parse(TermDate);        				
        			}
        			
        			else {
        				TermDate=JOptionPane.showInputDialog("Payresults already exist, PP EndDate:"+LastDate+"\n"
        						+ "Enter Termination Date - YYYY-MM-DD ");
        				TerminationDate=LocalDate.parse(TermDate);
        			}
        			
        			if(day.equals("SATURDAY")||day.equals("SUNDAY")) {
        				
        				TermDate=JOptionPane.showInputDialog("Employee cannot be terminated on weekend"+"\n"
        						+ "Enter Termination Date - YYYY-MM-DD ");
        				TerminationDate=LocalDate.parse(TermDate);
        			}
        			
        		}
        		
        		Cell cell2Update = EmployeeData.getRow(cell_Num).getCell(2);
        		cell2Update.setCellValue(TermDate);
        		
        		FileOutputStream outputStream = new FileOutputStream("Database.xls");
	            book.write(outputStream);
	            book.close();
	            outputStream.close();
	            
	            JOptionPane.showMessageDialog(null, "Employee Master Data is updated");
        	}
        	
        	if(ActionValue.equals("Change MasterData")) {
        		ChangeMasterDataTrigger=1;
        	ChangeMasterData change=new ChangeMasterData();
        	change.ChangeMasterDataMethod();	
        	}
        	}
        	
        	catch(NullPointerException e) {
        		JOptionPane.showMessageDialog(null, "Action Cancelled by user");
       	        System.exit(0);
       	}
        	catch(NumberFormatException e) {
        		JOptionPane.showMessageDialog(null, "Action Cancelled by user");
       	        System.exit(0);
        	}
        	
        	}
        
        else {
        	JOptionPane.showMessageDialog(null, "Action Cancelled by user");
        	}
        
        
        
        
        }
        
}

