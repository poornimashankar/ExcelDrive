import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	}

	public ArrayList<String> getData(String testCaseName) throws IOException {
		ArrayList<String> testDataList =  new ArrayList<String>();
		FileInputStream fis = new FileInputStream("C:\\Selenium\\Demo.xlsx");
		int k = 0;
		int column = 0;
		XSSFWorkbook workBook = new XSSFWorkbook(fis);
		int numOfSheets = workBook.getNumberOfSheets();
		Iterator rawIterator;
		XSSFSheet sheet;

		for (int i = 0; i < numOfSheets; i++) {

			if (workBook.getSheetName(i).equalsIgnoreCase("Testdata")) {
				sheet = workBook.getSheetAt(i);
				rawIterator = sheet.rowIterator();
				while (rawIterator.hasNext()) {
					Row firstRaw = (Row) rawIterator.next();
					Iterator<Cell> cellIterator = firstRaw.cellIterator();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						if (cell.getStringCellValue().equalsIgnoreCase("Testcases")) {
							column = k;
						}
						k++;
					}

				}

				System.out.println(column);
				rawIterator = sheet.rowIterator();
				Row r;
				while (rawIterator.hasNext()) {
					r = (Row) rawIterator.next();

					if (r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {

						Iterator<Cell> cv = r.cellIterator();

						while (cv.hasNext()) {
							Cell c =  cv.next();
							if(c.getCellType()==CellType.STRING) {
							testDataList.add(c.getStringCellValue());
							}
							else { 
								//testDataList.add(Double.toString(c.getNumericCellValue()));
								testDataList.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								
							}
						}
					}
				}

			}

		}
		return testDataList;
	}

}
