package task13;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) {

		ReadExcel obj = new ReadExcel();

		String value;

		try {

			obj.readingExcel();

			value = obj.readingExcel("Sheet1", 0, 0);

			System.out.println("Completed");

		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	public void readingExcel() throws Exception {

		String filePath = "C:\\Users\\Subasri Suresh\\OneDrive\\Desktop\\ReadExcel.xlsx";

		DataFormatter format = new DataFormatter();

		String result = null;

		FileInputStream inStream = new FileInputStream(filePath);

		XSSFWorkbook book = new XSSFWorkbook(inStream);

		for (int i = 1; i <= 5; i++) {

			for (int j = 1; j <= 5; j++) {

				XSSFCell cell = book.getSheet("Sheet1").getRow(i).getCell(j);

				result = format.formatCellValue(cell);

				System.out.print(result + " ");
			}

		}

		book.close();

	}

	public String readingExcel(String sheet, int row, int column) throws Exception {

		String result = null;

		String filePath = "C:\\Users\\Subasri Suresh\\OneDrive\\Desktop\\ReadExcel.xlsx";
		DataFormatter format = new DataFormatter();

		FileInputStream inStream = new FileInputStream(filePath);

		XSSFWorkbook book = new XSSFWorkbook(inStream);

		XSSFCell cell = book.getSheet(sheet).getRow(row).getCell(column);
		result = format.formatCellValue(cell);

		System.out.println(cell);

		book.close();

		return (result);
	}

}


/* Output


Name Age Email    
John Doe 30 john@test.com    
Jane Doe 28 john@test.com    
Bob Smith 35 jacky@example.com    
Swapnil 37 swapnil@example.com    
john@test.com                          */
