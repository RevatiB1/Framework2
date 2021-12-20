package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadNumerAsString {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		FileInputStream MyFile= new FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
//		String Value = WorkbookFactory.create(MyFile).getSheet("Sheet2").getRow(1).getCell(0).getStringCellValue();
//		
//		System.out.println(Value);
//		
		String Value1 = WorkbookFactory.create(MyFile).getSheet("nov expp").getRow(3).getCell(0).getStringCellValue();
	
		
		System.out.println(Value1);
	}

}