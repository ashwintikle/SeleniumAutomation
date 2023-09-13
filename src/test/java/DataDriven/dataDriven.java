package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public Object[][] getData(String sheetName) throws IOException {

		DataFormatter dataFormatter = new DataFormatter();
		XSSFSheet worksheet = null;
		FileInputStream fileInputStream = new FileInputStream("D:\\Ashwin_Docs\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		int noOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < noOfSheets; i++) {
			if (workbook.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName)) {
				worksheet = workbook.getSheetAt(i);
			}
		}
		int rowCount = worksheet.getPhysicalNumberOfRows();
		Row firstRow = worksheet.getRow(0);
		int columnCount = firstRow.getLastCellNum();
		Object[][] testData = new Object[rowCount - 1][columnCount];

		for (int i = 0; i <= rowCount - 1; i++) {

			Row testDataRow = worksheet.getRow(i + 1);// Fetch the second row as the first row has headers

			for (int j = 0; j <= columnCount; j++) {

				Cell testDataCell = testDataRow.getCell(j);
				testData[i][j] = dataFormatter.formatCellValue(testDataCell);
			}
		}

		return testData;
	}

}
