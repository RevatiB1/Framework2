package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadAllDataFromConstantColumn {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream MyFile= new FileInputStream("C:\\Users\\ADMIN\\OneDrive\\Desktop\\Marks'.xlsx");
		Sheet sh = WorkbookFactory.create(MyFile).getSheet("Students");
		
		int RowIndex = sh.getLastRowNum();
		
		for(int i=0;i<=RowIndex;i++)
			
		{
			String Value = sh.getRow(i).getCell(0).getStringCellValue();
			System.out.println(Value);
		}

	}

}