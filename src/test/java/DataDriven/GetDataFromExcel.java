package DataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetDataFromExcel {

	public ArrayList<String> getData(String scenarioName) throws IOException {

		ArrayList<String> data = new ArrayList<String>();
		FileInputStream fileInputStream = new FileInputStream("D:\\Ashwin_Docs\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet worksheet = null;

		int countOfSheets = workbook.getNumberOfSheets();
		for (int i = 0; i < countOfSheets; i++) {
			if (workbook.getSheetName(i).equals("SubmitOrderData")) {
				worksheet = workbook.getSheetAt(i);
			}
		}
		Iterator<Row> rows = worksheet.rowIterator();
		Row firstRow = rows.next();
		Iterator<Cell> cellItr = firstRow.cellIterator();
		int coloumn = 0;
		int k = 0;
		while (cellItr.hasNext()) {
			Cell cellValue = cellItr.next();
			if (cellValue.getStringCellValue().equalsIgnoreCase("Scenario")) {
				coloumn = k;
			}
			k++;
		}
		while (rows.hasNext()) {
			Row successfulOrder = rows.next();
			if (successfulOrder.getCell(coloumn).getStringCellValue().equalsIgnoreCase(scenarioName)) {
				Iterator<Cell> getSuccessOrderTestDataRow = successfulOrder.cellIterator();
				while (getSuccessOrderTestDataRow.hasNext()) {
					Cell successOrderTestDataCell = getSuccessOrderTestDataRow.next();
					if (successOrderTestDataCell.getCellType() == CellType.STRING) {
						data.add(successOrderTestDataCell.getStringCellValue());
					} else {
						data.add(NumberToTextConverter.toText(successOrderTestDataCell.getNumericCellValue()));
					}
				}
			}
		}
		return data;
	}
}
