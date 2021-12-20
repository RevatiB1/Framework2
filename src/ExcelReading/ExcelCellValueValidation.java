package ExcelReading;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class ExcelCellValueValidation {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		FileInputStream MyFile= new FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
		//CellType CT = WorkbookFactory.create(MyFile).getSheet("Sheet1").getRow(0).getCell(0).getCellType();
		
		//System.out.println(CT);
		
//		CellType CT1 = WorkbookFactory.create(MyFile).getSheet("Sheet1").getRow(1).getCell(0).getCellType();
//		
//		System.out.println(CT1);
		
		CellType CT2 = WorkbookFactory.create(MyFile).getSheet("nov expp").getRow(2).getCell(0).getCellType();
	
		System.out.println(CT2);
	
	}

}