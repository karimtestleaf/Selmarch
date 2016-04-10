import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataa {
	public static void main(String[] args) throws IOException {
		XSSFWorkbook wbook = new XSSFWorkbook();
		XSSFSheet sheet = wbook.createSheet("TestResult");

		// Create Row objects
		XSSFRow row1 = sheet.createRow(0);
		XSSFRow row2 = sheet.createRow(1);
		XSSFRow row3 = sheet.createRow(2);
		XSSFRow row4 = sheet.createRow(3);

		// Header
		row1.createCell(0).setCellValue("SNo");
		row1.createCell(1).setCellValue("testcase");
		row1.createCell(2).setCellValue("status");

		// Create test names
		row2.createCell(1).setCellValue("create");
		row3.createCell(1).setCellValue("delete");
		row4.createCell(1).setCellValue("merge");

		for (int i = 1; i < 3; i++) {
			if (i % 2 == 0) {
				row4.createCell(2).setCellValue("pass");
				row2.createCell(2).setCellValue("pass");
			} else {
				row3.createCell(2).setCellValue("fail");
			}
		}

		// Create this when you need to write / update
		FileOutputStream fileOutput = new FileOutputStream(new File("C:\\Sheets\\output1.xlsx"));
		wbook.write(fileOutput);
		fileOutput.close();

	}

}
