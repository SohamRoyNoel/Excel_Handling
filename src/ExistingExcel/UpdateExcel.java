package ExistingExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcel {

	public static void main(String[] args) {
		XSSFWorkbook workbook=null;
		XSSFSheet sheet;
		try{
			FileInputStream file = new FileInputStream(new File("E:\\HarTest\\TargetExcel\\Targets.xlsx"));

			//Create Workbook instance 
			workbook = new XSSFWorkbook(file);

		
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

			Employee ess = new Employee(6,"John","Cena");
			//Get the count in sheet
			int rowCount = sheet.getLastRowNum()+1;
			Row empRow = sheet.createRow(rowCount);
			System.out.println();
			Cell c1 = empRow.createCell(0);
			c1.setCellValue(ess.getId());
			Cell c2 = empRow.createCell(1);
			c2.setCellValue(ess.getFirstName());
			Cell c3 = empRow.createCell(2);
			c3.setCellValue(ess.getLastName());
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		try
		{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new 
					File("E:\\HarTest\\TargetExcel\\Targets.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("Updated");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}

}
