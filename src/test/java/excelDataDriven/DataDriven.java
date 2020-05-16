package excelDataDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class DataDriven {
	
	public ArrayList<String> getData(String testcaseName , String SheetName) throws IOException
	{
		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream("C:\\Users\\gole\\Desktop\\demoData.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		int sheets = workbook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) 
		{
			if (workbook.getSheetName(i).equalsIgnoreCase(SheetName)) 
			{
				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator<Row> rows = sheet.iterator(); // collection of rows
				Row firstrow = rows.next();

				Iterator<Cell> cell = firstrow.cellIterator();// collection of cell
				
				int k = 0;
				int col = 0;
				while (cell.hasNext()) 
				{
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{
						col = k;
					}
					k++;
				}
				System.out.println(col);

				while (rows.hasNext())
				{
					Row r = rows.next();
					if (r.getCell(col).getStringCellValue().equalsIgnoreCase(testcaseName)) 
					{
						Iterator<Cell> Cell1 = r.cellIterator();
						
						while (Cell1.hasNext()) 
						{
							Cell c = Cell1.next();
							if(c.getCellTypeEnum()==CellType.STRING)
							{
								a.add(c.getStringCellValue());
							}
							else
							{
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
							
						}
					}
				}
			}
		}
		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
					}
		}
	
