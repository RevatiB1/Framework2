
package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelRowCount {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream MyFile= new FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
		
		int TotalRows = WorkbookFactory.create(MyFile).getSheet("nov expp").getLastRowNum()+1;
	
		System.out.println(TotalRows);
		

	}

}