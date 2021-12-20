package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadAllDatatypesCombinationOfRowHavingConstantColumn {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		FileInputStream MyFile= new FileInputStream("C:\\Users\\ADMIN\\OneDrive\\Desktop\\Marks'.xlsx");
		Sheet sh = WorkbookFactory.create(MyFile).getSheet("Students");
		
		//calculating last row to get count for traversing.
		int LastRowIndex = sh.getLastRowNum();
		System.out.println("Traversing till row index number "+LastRowIndex);
		
		//Traversing and trying to get datatype of each cell value and then using if else to print data
		
		for(int i=0; i<=LastRowIndex;i++)
		{
			Cell CellInfo = sh.getRow(i).getCell(0);
			//cell info will hold value from respective cell
			
			CellType CellDataType = CellInfo.getCellType();//numeric,string,boolean
			//Celldatatype found using getcelltype method
			
			if(CellDataType==CellType.STRING)
			{
				String Value = CellInfo.getStringCellValue();
				System.out.println(Value);
				
			}
			else if (CellDataType==CellType.NUMERIC) 
			{
			double Value = CellInfo.getNumericCellValue();	
			System.out.println(Value);
			}
			else if (CellDataType==CellType.BOOLEAN) 
			{
			boolean Value = CellInfo.getBooleanCellValue();	
			System.out.println(Value);
			}
		}

	}

}