package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelDoubleValueReading {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
		//Step1. Create File using FileInputStream --> Give filePath along with file name and extension
		FileInputStream MyFile= new FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
		//step2.
		
		double Value = WorkbookFactory.create(MyFile).getSheet("nov expp").getRow(1).getCell(0).getNumericCellValue();
		
		System.out.println(Value);
		
		
		

	}

}