package task13;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		WriteExcel obj = new WriteExcel();
		
		try {
			obj.createExcel();
			System.out.println(" ");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
     }
	
	public void createExcel() throws Exception {
		
		String filePath = "C:\\Users\\Subasri Suresh\\OneDrive\\Desktop\\WriteExcel.xlsx";
		
		File file = new File(filePath);
		
		FileOutputStream outStream = new FileOutputStream(file);
		
		XSSFWorkbook book = new XSSFWorkbook();
		
		XSSFSheet sheet = book.createSheet("Sheet1");
		
		sheet.createRow(0).createCell(0).setCellValue("Name");
		sheet.createRow(0).createCell(1).setCellValue("Age");
		sheet.createRow(0).createCell(2).setCellValue("Email");
		
		sheet.createRow(1).createCell(0).setCellValue("John Doe");
		sheet.getRow(1).createCell(1).setCellValue("30");
		sheet.getRow(1).createCell(2).setCellValue("john@test.com");

		sheet.createRow(2).createCell(0).setCellValue("Jane Doe");
		sheet.getRow(2).createCell(1).setCellValue("28");
		sheet.getRow(2).createCell(2).setCellValue("john@test.com");

		sheet.createRow(3).createCell(0).setCellValue("Bob Smith");
		sheet.getRow(3).createCell(1).setCellValue("35");
		sheet.getRow(3).createCell(2).setCellValue("jacky@example.com");

		sheet.createRow(4).createCell(0).setCellValue("Swapnil ");
		sheet.getRow(4).createCell(1).setCellValue("37");
		sheet.getRow(4).createCell(2).setCellValue("joe@example.com");

		book.write(outStream);
		
		book.close();
		outStream.close();
		
	}
}

/*

Output

Name	    Age	Email
John Doe	30	john@test.com
Jane Doe	28	john@test.com
Bob Smith	35	jacky@example.com
Swapnil 	37	joe@example.com       */


