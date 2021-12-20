package ExcelReading;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadFullExcel {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
	
		FileInputStream MyFile= new FileInputStream("C:\\Users\\ADMIN\\OneDrive\\Desktop\\Marks'.xlsx");
		Sheet sh = WorkbookFactory.create(MyFile).getSheet("Students");
		
		int LastRowIndex = sh.getLastRowNum();
		//System.out.println(LastRowIndex);
		
		for(int i=0;i<=LastRowIndex;i++)//0,1--> for row reading
		{
			
			int LastColumnIndex = sh.getRow(i).getLastCellNum()-1;//3
			//getrow=row datatype
			for(int j=0;j<=LastColumnIndex;j++)//0,1,2,3--> for column reading
			{
				Cell CellInfo = sh.getRow(i).getCell(j);
				//cell info will hold value from respective cell
				
				CellType CellDataType = CellInfo.getCellType();//numeric,string,boolean
				//Celldatatype found using getcelltype method
				
				if(CellDataType==CellType.STRING)
				{
					String Value = CellInfo.getStringCellValue();
					System.out.print(Value);
					
				}
				else if (CellDataType==CellType.NUMERIC) 
				{
				double Value = CellInfo.getNumericCellValue();	
				System.out.print(Value);
				}
				else if (CellDataType==CellType.BOOLEAN) 
				{
				boolean Value = CellInfo.getBooleanCellValue();	
				System.out.print(Value);
				}
			
			}
			System.out.println();
			
			
		}
		

	}

}