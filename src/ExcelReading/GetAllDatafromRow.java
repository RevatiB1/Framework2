package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class GetAllDatafromRow {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream MyFile= new FileInputStream("F:\\Downloads\\akshay audio\\hishob.xlsx");
		Sheet sh = WorkbookFactory.create(MyFile).getSheet("Sheet2");
		
		//static type
//		for(int i=0;i<=3;i++)//0,1
//		{
//			String Value = sh.getRow(0).getCell(i).getStringCellValue();
//			System.out.println(Value);
//		}
		
		//Dynamic coding for Row reading
		
		int LastCellIndex = sh.getRow(1).getLastCellNum()-1;
		
		System.out.println("Cell index count is "+LastCellIndex);
		
		
		for(int i=0; i<=LastCellIndex;i++)
		{
			String Value = sh.getRow(0).getCell(i).getStringCellValue();
			System.out.println(Value);
		}
		
	}

}