package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadAllTypeDatafromFirstCell {

	private static final CellType STRING = null;

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream MyFile= new FileInputStream("C:\\Users\\ADMIN\\OneDrive\\Desktop\\Marks'.xlsx");
		//create object by passing path
		Sheet sh = WorkbookFactory.create(MyFile).getSheet("Students");
	//	we will go to sheet
		Cell CellInfo = sh.getRow(0).getCell(0);
		System.out.println("Cell::"+CellInfo);//It returns cell as roll no
//		cell cha type bhgnr
		CellType CellInfoType = CellInfo.getCellType();//this checks for info datatype
		//Value kontya type chi ahe he bhgnr
		System.out.println(CellInfoType);
	//we r chec	
		if(CellInfoType==CellType.STRING)
		{
			
			String value = CellInfo.getStringCellValue();
			System.out.println(value);
		}
		
		else if (CellInfoType==CellType.NUMERIC) 
		{
			double value = CellInfo.getNumericCellValue();
			System.out.println(value);
		}
		
		
		else if (CellInfoType==CellType.BOOLEAN) 
		{
		
			boolean value = CellInfo.getBooleanCellValue();
			System.out.println(value);
		}
		
		

	}

}