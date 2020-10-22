import javax.swing.*;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeCheck {

public static String EmployeeID;
public static int CellID;
	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		
	try {	
		EmployeeNumericCheck EmployeeCheck= new EmployeeNumericCheck();
		int a= EmployeeCheck.NumericValidation();
			while(a!=1) {
				JOptionPane.showMessageDialog(null, "Employee ID should be numeric and 5 digits");
				a = EmployeeCheck.NumericValidation();
			}
		EmployeeID=EmployeeNumericCheck.EmployeeID1;
	}
	 catch(NullPointerException e) 
    { 
		 JOptionPane.showMessageDialog(null, "Input cancelled by user");
        System.exit(0);
    } 
	
	try {
		InputStream myxls = new FileInputStream("Database.xls");
		Workbook book = new HSSFWorkbook(myxls);
		org.apache.poi.ss.usermodel.Sheet sheet = book.getSheetAt(0);
		int LastRowNumber= sheet.getLastRowNum();
		
		int found=0;
		
		for(int i=1;i<=LastRowNumber;i++) 
		{
			Cell cell2= book.getSheetAt(0).getRow(i).getCell(0); 
            String output=cell2.toString();
            
            float output2=Float.parseFloat(output);
            
			if(output.contentEquals(EmployeeID)) 
			{
				found=found+1;
				CellID=i;
				break;
			}       
		}
		myxls.close();
		if(found==1) 
		{
			JOptionPane.showMessageDialog(null, "Existing EE");
			
			ExistingEmployee existing = new ExistingEmployee();
			existing.ExistingEmployeeMethod();
			/*int confirm = JOptionPane.showConfirmDialog(null, "Run Payroll?", null, JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE );
			if(confirm==0) 
			{	
				PayrollRun NewHire=new PayrollRun();
				NewHire.PayrollRunMethod();
			}
			else
			{
				JOptionPane.showMessageDialog(null, "Payroll is not run");
				System.exit(0);
			}*/
		}
		
		else
		{
			int confirm = JOptionPane.showConfirmDialog(null, "New EE, hire?", EmployeeID, JOptionPane.OK_CANCEL_OPTION, JOptionPane.INFORMATION_MESSAGE );
				if(confirm==0) 
				{	CellID=LastRowNumber+1;
					NewHire newhire = new NewHire();
					((NewHire) newhire).NewHireMethod();
				}
				else
				{
					JOptionPane.showMessageDialog(null, "Hiring cancelled by user");
					System.exit(0);
				}
		}
	}
	catch(FileNotFoundException e) {
		 System.out.print("Database is corrupted");
	        System.exit(0);
	}
	}

}
