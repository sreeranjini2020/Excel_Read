package excel.read;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static XSSFWorkbook w;



	public static XSSFSheet s;



	public static FileInputStream f;



	public static String readStringData(int i,int j) throws IOException {



	f= new FileInputStream("C:\\Users\\minnu\\git\\Excel_Read\\Excel_Read\\src\\main\\resources\\Student.xlsx");



	w= new XSSFWorkbook(f);



	s= w.getSheet("Sheet1");

	Row r=s.getRow(i);



	Cell c=r.getCell(j);



	return c.getStringCellValue();// string value will read



	}



	public static String readIntegerData(int i,int j) throws IOException {

		



			f= new FileInputStream("C:\\Users\\minnu\\git\\Excel_Read\\Excel_Read\\src\\main\\resources\\Student.xlsx");



			w= new XSSFWorkbook(f);



			s= w.getSheet("Sheet1");



			Row r=s.getRow(i);



			Cell c=r.getCell(j);



			int value=(int) c.getNumericCellValue();//numberic cell value read(it will be integer, double, float) int is used to type cast



			return String.valueOf(value); //return type is string so given String



			}

}
