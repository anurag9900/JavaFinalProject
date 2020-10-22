import java.util.Arrays;
import java.util.List;
import wagu.Block;
import wagu.Board;
import wagu.Table;
import javax.swing.*;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.Math;
import java.text.DecimalFormat;
import java.util.*;
import java.time.*;
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

public class DisplayPaySlips {

	private static DecimalFormat df = new DecimalFormat("0.00");
    static LocalDate DateDisplay=LocalDate.now();
    static LocalTime TimeDisplay=LocalTime.now();
    
    public void GeneratePaySlip(String PayPeriod) throws EncryptedDocumentException, InvalidFormatException, IOException {
    	InputStream inputStream= new FileInputStream(EmployeeCheck.EmployeeID+".xls");
		//Workbook book1 = new HSSFWorkbook(myxls1);
		//org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0);
		Workbook book1 = WorkbookFactory.create(inputStream);
        org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0);                
        int CellID=0;
		String[] PaySlipOutput=new String[17];
		for(int i=1;i<=PayResults.getLastRowNum();i++) {
			Cell cell2= PayResults.getRow(i).getCell(0); 
			if(cell2.toString().equals(PayPeriod)) {
				CellID=i;
				break;
			}
		}
		
		for(int i=0;i<=16;i++) {
			PaySlipOutput[i]=book1.getSheetAt(0).getRow(CellID).getCell(i).toString(); 			
		}
        
		float overtime_hours=(float) (Float.parseFloat(PaySlipOutput[12])*1.5);
        String OH=df.format(overtime_hours);
        float TotalGross=(float)(Float.parseFloat(PaySlipOutput[3])+Float.parseFloat(PaySlipOutput[4]));
        String TotalGross_Update=df.format(TotalGross);
        
        
        String company = ""
                + "Payroll Processing Team Ltd\n"
                + "111 Yorkland Blvd #400, North York, ON M2J AAA\n"
                + " \n"
                + "PAY SLIP"
                + " \n";
        List<String> t1Headers = Arrays.asList("INFO", "Employee");
        List<List<String>> t1Rows = Arrays.asList(
                Arrays.asList("Date: "+DateDisplay.toString(), "Employee ID: "+EmployeeCheck.EmployeeID),
                Arrays.asList("Time: "+TimeDisplay.toString().substring(0, 5), ""),
                Arrays.asList("PP Num: "+PaySlipOutput[0], "PP SDate: "+ PaySlipOutput[1]),
                Arrays.asList("", "PP EDate: "+ PaySlipOutput[2])
        );
        String t2Desc = "PAYMENT DETAILS";
        List<String> t2Headers = Arrays.asList("Payments", "Hourly Rate", "Hours", "Amount");
        List<List<String>> t2Rows = Arrays.asList(
                Arrays.asList("Gross", PaySlipOutput[12],PaySlipOutput[13] , PaySlipOutput[3]),
                Arrays.asList("OverTime", OH,PaySlipOutput[14], PaySlipOutput[4]),
                Arrays.asList("RetroAmount", " "," ", PaySlipOutput[16])
        );
        List<Integer> t2ColWidths = Arrays.asList(17, 8, 6, 12);
        String t3Desc = "DEDUCTION Details";
        List<String> t3Headers = Arrays.asList("Deductions", "Amount");
        List<List<String>> t3Rows = Arrays.asList(
                Arrays.asList("Fed Tax", PaySlipOutput[5]),
                Arrays.asList("CPP", PaySlipOutput[6]),
                Arrays.asList("EI", PaySlipOutput[7]),
                Arrays.asList("Garnishment", PaySlipOutput[8]),
                Arrays.asList("Leaves", PaySlipOutput[15])
        );
        String summary = ""
                + "GROSS\n"
                + "Total Deductions\n"
                + "Net Payment\n"
                ;
        String summaryVal = ""
                + TotalGross_Update+"\n"
                + PaySlipOutput[9]+"\n"
                + PaySlipOutput[11]+"\n"
                ;
        

        //bookmark
        Board b = new Board(48);
        b.setInitialBlock(new Block(b, 46, 7, company).allowGrid(false).setBlockAlign(Block.BLOCK_CENTRE).setDataAlign(Block.DATA_CENTER));
        b.appendTableTo(0, Board.APPEND_BELOW, new Table(b, 48, t1Headers, t1Rows));
        b.getBlock(3).setBelowBlock(new Block(b, 46, 1, t2Desc).setDataAlign(Block.DATA_CENTER));
        b.appendTableTo(5, Board.APPEND_BELOW, new Table(b, 48, t2Headers, t2Rows, t2ColWidths));
        b.getBlock(10).setBelowBlock(new Block(b, 46, 1, t3Desc).setDataAlign(Block.DATA_CENTER));
        b.appendTableTo(14, Board.APPEND_BELOW, new Table(b, 48, t3Headers, t3Rows));
        Block summaryBlock = new Block(b, 35, 9, summary).allowGrid(false).setDataAlign(Block.DATA_MIDDLE_RIGHT);
        b.getBlock(17).setBelowBlock(summaryBlock);
        Block summaryValBlock = new Block(b, 12, 9, summaryVal).allowGrid(false).setDataAlign(Block.DATA_MIDDLE_RIGHT);
        summaryBlock.setRightBlock(summaryValBlock);
        
        
        //b.showBlockIndex(true);
        System.out.println(b.invalidate().build().getPreview());

    }

}