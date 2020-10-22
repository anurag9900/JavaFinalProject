import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.ChronoField;
import java.time.temporal.ChronoUnit;

import javax.swing.JOptionPane;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PayrollRun {
	
	private static DecimalFormat df = new DecimalFormat("0.00");
	@SuppressWarnings("unlikely-arg-type")
	public void PayrollRunMethod() throws IOException, EncryptedDocumentException, InvalidFormatException {
		
		InputStream myxls = new FileInputStream("Database.xls");
		Workbook book = new HSSFWorkbook(myxls);
		org.apache.poi.ss.usermodel.Sheet EmployeeData = book.getSheetAt(0);
		org.apache.poi.ss.usermodel.Sheet PayPeriod = book.getSheetAt(1);
		org.apache.poi.ss.usermodel.Sheet EmployeeDataChange = book.getSheet("Change Data");
		
		InputStream inputStream= new FileInputStream(EmployeeCheck.EmployeeID+".xls");
		//Workbook book1 = new HSSFWorkbook(myxls1);
		//org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0);
		Workbook book1 = WorkbookFactory.create(inputStream);
        org.apache.poi.ss.usermodel.Sheet PayResults = book1.getSheetAt(0);
		
		
		
		int CellID=EmployeeCheck.CellID;
		int PP=0;
		int end=0,start=0;
		LocalDate PayrollStartDate, PayrollEndDate;
		float retroAmount=0;
		
		String[] EmployeeDataRow= new String[14];
		for(int i=1;i<=13;i++) {
		EmployeeDataRow[i]= EmployeeData.getRow(CellID).getCell(i).toString();
		}
		String[] EmployeeDataChangeRow= new String[10];
		for(int i=1;i<=EmployeeDataChange.getLastRowNum();i++) 
		{
			Cell cell2= book.getSheet("Change Data").getRow(i).getCell(0); 
            String output=cell2.toString();            
			if(output.contentEquals(EmployeeCheck.EmployeeID)) 
			{
				for(int j=0;j<=9;j++) {
					EmployeeDataChangeRow[j]= EmployeeDataChange.getRow(i).getCell(j).toString();
				}
			}       
		}
		
		
		String[] PayPeriodNumber = new String[PayPeriod.getLastRowNum()+1];
		String [] PayPeriodStartDate = new String[PayPeriod.getLastRowNum()+1];
		String [] PayPeriodEndDate = new String[PayPeriod.getLastRowNum()+1];
		for(int i=1;i<=PayPeriod.getLastRowNum();i++) {
			PayPeriodNumber[i]=PayPeriod.getRow(i).getCell(0).toString();
			PayPeriodStartDate[i]=PayPeriod.getRow(i).getCell(1).toString();
			PayPeriodEndDate[i]=PayPeriod.getRow(i).getCell(2).toString();
		}
		
		
		LocalDate HireDate=LocalDate.parse(EmployeeDataRow[1]);
		LocalDate TermDate=LocalDate.parse(EmployeeDataRow[2]);
		//if(PayResults.getLastRowNum()!=0) 
		//{	
			int PP_Num=0;String output="Enter Pay Period for Payroll Run\n";
			if(PayResults.getLastRowNum()!=0) 
				{
			//PP_Num=Integer.parseInt(PayResults.getRow(PayResults.getLastRowNum()).getCell(0).toString().substring(0,2).replace(".",""));
				PP_Num=Integer.parseInt(PayResults.getRow(PayResults.getLastRowNum()).getCell(0).toString());
				}
			for(int i=PP_Num+1;i<=12;i++) 
						{
				LocalDate Start=LocalDate.parse(PayPeriodStartDate[i]);
				LocalDate End=LocalDate.parse(PayPeriodEndDate[i]);
				if(HireDate.isAfter(Start) && HireDate.isBefore(End)) {
					start=i;
							}
				if(TermDate.isAfter(Start)&&TermDate.isBefore(End)) {
					end=i;
							}				
						}
			if(end==0) end=12;
			if(start==0) start=PP_Num;
			
			if(PayResults.getLastRowNum()!=0) {start=start+1;}
			
					for(int j=start;j<=end;j++) {
				output=output+ "Period:"+PayPeriodNumber[j].substring(0,2).replace(".","")+"   "+
						"Start Date:"+ PayPeriodStartDate[j]+"   "+
						"End Date:"+PayPeriodEndDate[j]+"\n";
						}				
			PP=Integer.parseInt(JOptionPane.showInputDialog(output));
			int PP1=Integer.parseInt(PayPeriodNumber[start].substring(0,2).replace(".",""));
			
			while(PP<PP1 && PP>12)
			{String output1="Enter Correct Number"+"\n"+output;
				PP=Integer.parseInt(JOptionPane.showInputDialog(output1));
			}
			
			String PP_Update=Integer.toString(PP);
			if(start!=0 && start!=PP_Num+1) PayrollStartDate=HireDate; else PayrollStartDate=LocalDate.parse(PayPeriodStartDate[PP_Num+1]);
			if(end!=12 && end!=0) PayrollEndDate=TermDate; else PayrollEndDate=LocalDate.parse(PayPeriodEndDate[PP]);
			
			float[] PayResults1= PayrollRunRetro(PayrollStartDate, PayrollEndDate, EmployeeDataRow,1);
			
			if(EmployeeDataRow[11].equals("1")) {
				
			float[] PayResults2= PayrollRunRetro(LocalDate.parse(EmployeeDataChangeRow[1]), LocalDate.parse(EmployeeDataChangeRow[2]), EmployeeDataRow,0);
			float[] PayResults3= PayrollRunRetro(LocalDate.parse(EmployeeDataChangeRow[1]), LocalDate.parse(EmployeeDataChangeRow[2]), EmployeeDataChangeRow,0);
			retroAmount=PayResults2[7]-PayResults3[7];
			int cell_Num= EmployeeCheck.CellID;
    		Cell cell2Update = EmployeeData.getRow(cell_Num).getCell(11);
    		cell2Update.setCellValue("0");
			}
			
			
			String PayrollStartDate_Update=PayrollStartDate.toString();
			String PayrollEndDate_Update=PayrollEndDate.toString();
			String GrossPay_Update=df.format(PayResults1[0]);
			String OverTimePay_Update=df.format(PayResults1[1]);
			String FedTax_Update=df.format(PayResults1[2]);
			String CPP_Update=df.format(PayResults1[3]);
			String EI_Update=df.format(PayResults1[4]);
			String Garnishment_Update=df.format(PayResults1[5]);
			String TotalDeductions_Update=df.format(PayResults1[6]);
			String NetPay_Udpate=df.format(PayResults1[7]-retroAmount);
			String HourlyRate_update=df.format(PayResults1[8]);
			String TotalHours_Update=df.format(PayResults1[9]);
			String OverTimeHours=EmployeeDataRow[8];
			String Leaves_Update=df.format(PayResults1[10]);
			String RetroAmount=df.format(retroAmount);
			
			String[] PayResultsUpdate = {PP_Update,PayrollStartDate_Update,PayrollEndDate_Update,
					GrossPay_Update,OverTimePay_Update,FedTax_Update,CPP_Update,EI_Update
					,Garnishment_Update,TotalDeductions_Update, EI_Update, NetPay_Udpate, 
					HourlyRate_update, TotalHours_Update, OverTimeHours,Leaves_Update, RetroAmount};
			
			int CellID1=PayResults.getLastRowNum();
			Row row = PayResults.createRow(CellID1+1);
			
			for(int i=0; i<=16;i++ ) {
				row.createCell(i).setCellValue(PayResultsUpdate[i]);
			}
          
			int cell_Num= EmployeeCheck.CellID;
    		Cell cell2Update = EmployeeData.getRow(cell_Num).getCell(13);
    		cell2Update.setCellValue("0");
    		
    		FileOutputStream outputStream1 = new FileOutputStream("Database.xls");
            book.write(outputStream1);
            book.close();
            outputStream1.close();
			
			inputStream.close();
			
			//Modifying the EmployeeData Payresults file
			 FileOutputStream outputStream = new FileOutputStream(EmployeeCheck.EmployeeID+".xls");
	            book1.write(outputStream);
	            book1.close();
	            outputStream.close();
			
			
			DisplayPaySlips display=new DisplayPaySlips();
			display.GeneratePaySlip(PP_Update);
		
	}
	
	public float[] PayrollRunRetro(LocalDate PayrollStartDate, LocalDate PayrollEndDate,String[] EmployeeDataRow, int Trigger ) {
		
		float NoOfWorkingDays=0,TotalDeductions=0, leaves=0, TotalHours=0,Garnishment=0,HourlyRate=0, FedTax=0, EI=0, CPP=0,GrossPay=0, HoursPerDay=0,NetPay=0, OverTimePay=0;
		
		String[][] province= {
				{"AB","0.12"}, {"BC","0.11"}, {"MB","0.15"}, {"NB","0.7"}, {"NF","0.10"}, {"NS","0.12"}, 
				{"NT","0.11"}, {"ON","0.16"}, {"PE","0.8"}, {"QC","0.15"},{"SK","0.10"}, {"YT","0.7"}, {"NN","0.8"}	
		};
		float province_rate = 0;
		for(String x[]:province) {
			if(EmployeeDataRow[4].equals(x[0])) {
				province_rate=Float.parseFloat(x[1]);
			}
		}
		
		long NumberOfCalendarDays=ChronoUnit.DAYS.between(PayrollStartDate, PayrollEndDate)+1;
		NoOfWorkingDays=NumberOfCalendarDays;
		
		// calculate number of working days
		for(long i=1;i<=NumberOfCalendarDays;i++) {
			String day=DayOfWeek.of(PayrollStartDate.plusDays(i).get(ChronoField.DAY_OF_WEEK)).toString();
			if((day.equals("SUNDAY")) ||(day.equals("SATURDAY"))) {
				NoOfWorkingDays=NoOfWorkingDays-1;
			}
		}
		
		//check hourly rate
		if(EmployeeDataRow[5].equals("S")) {
		HourlyRate=(Float.parseFloat(EmployeeDataRow[6]))/((Float.parseFloat(EmployeeDataRow[7])*52)/12);
		}
		else HourlyRate=Float.parseFloat(EmployeeDataRow[7]);
		
		//calculation of hours per day
		HoursPerDay= Float.parseFloat(EmployeeDataRow[7])/5;
		
		//calculation of Gross Pay			
		GrossPay=NoOfWorkingDays*HourlyRate*HoursPerDay;
		
		//Calculation of Over Time
		OverTimePay=(float) (Float.parseFloat(EmployeeDataRow[8])*(HourlyRate*1.5));
		
		//overtime hours
		
		//Garnishment
		
		Garnishment=Float.parseFloat(EmployeeDataRow[9]);
		
		//Fed Tax
		FedTax=(float) (GrossPay*province_rate);
		
		//CPP
		CPP=(float) (GrossPay*0.05);
		
		//EI
		EI=(float) (GrossPay*.10);
		
		//
		TotalDeductions=FedTax+CPP+EI+Garnishment;		
		
		//total hours
		TotalHours=NoOfWorkingDays*HoursPerDay;
		
		
		//leave amount
		if(Trigger==1) {
		leaves=HourlyRate*HoursPerDay*(Float.parseFloat(EmployeeDataRow[13]));
		}
		
		//Calculation of Net Pay
		NetPay= GrossPay-Garnishment+OverTimePay-FedTax-CPP-EI-leaves;
		float[] PayResults= {GrossPay, OverTimePay, FedTax, CPP, EI,
				Garnishment, TotalDeductions, NetPay, HourlyRate, TotalHours,leaves};
		
		return PayResults;
		
	}
	
}
