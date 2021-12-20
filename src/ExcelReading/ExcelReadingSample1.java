package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReadingSample1 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
FileInputStream Myfile= new  FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
//String OutPut = WorkbookFactory.create(Myfile).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
//		
//System.out.println(OutPut);

String NovSheet = WorkbookFactory.create(Myfile).getSheet("nov expp").getRow(1).getCell(1).getStringCellValue();
	
System.out.println(NovSheet);


//Workbook book = WorkbookFactory.create(Myfile);
//Sheet sh = book.getSheet("Sheet1");
//Row rw = sh.getRow(0);
//Cell cl = rw.getCell(0);
//String value = cl.getStringCellValue();
//System.out.println(value);


		
	}

}