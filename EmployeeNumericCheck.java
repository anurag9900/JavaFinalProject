import javax.swing.JOptionPane;

public class EmployeeNumericCheck {
	public static String EmployeeID1;
	public int NumericValidation(){
		int z=0;
		String EmployeeIDcheck=JOptionPane.showInputDialog("Enter Employee ID (Numeric-5 Digits) ");
		int length=EmployeeIDcheck.length();
	if(EmployeeIDcheck.matches("[0-9]+")&& (length==5))
		{
		z=1;
		EmployeeID1=EmployeeIDcheck;
		}
	return z;
}
}