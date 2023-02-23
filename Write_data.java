package datadriven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_data {
	
	public static void main(String[] args) throws IOException {
		
		File f=new File("C:\\Users\\Admin\\eclipse-workspace\\Maven_Test\\datadriven.xlsx");
		FileInputStream fi=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(fi);
		w.createSheet("Sheet4").createRow(0).createCell(0).setCellValue("ID");
		w.getSheet("Sheet4").getRow(0).createCell(1).setCellValue("NAME");
		w.getSheet("Sheet4").getRow(0).createCell(2).setCellValue("DEPARTMENT");
		w.getSheet("Sheet4").getRow(0).createCell(3).setCellValue("ADDRESS");
		
		w.getSheet("Sheet4").createRow(1).createCell(0).setCellValue("1");
		w.getSheet("Sheet4").getRow(1).createCell(1).setCellValue("Mayil");
		w.getSheet("Sheet4").getRow(1).createCell(2).setCellValue("Cse");
		w.getSheet("Sheet4").getRow(1).createCell(3).setCellValue("Chennai");
		
		w.getSheet("Sheet4").createRow(2).createCell(0).setCellValue("2");
		w.getSheet("Sheet4").getRow(2).createCell(1).setCellValue("Sibu");
		w.getSheet("Sheet4").getRow(2).createCell(2).setCellValue("Ece");
		w.getSheet("Sheet4").getRow(2).createCell(3).setCellValue("Dharmapuri");
		
		w.getSheet("Sheet4").createRow(3).createCell(0).setCellValue("3");
		w.getSheet("Sheet4").getRow(3).createCell(1).setCellValue("Jai");
		w.getSheet("Sheet4").getRow(3).createCell(2).setCellValue("Mech");
		w.getSheet("Sheet4").getRow(3).createCell(3).setCellValue("Bangalore");
		
		w.getSheet("Sheet4").createRow(4).createCell(0).setCellValue("4");
		w.getSheet("Sheet4").getRow(4).createCell(1).setCellValue("Thaaru");
		w.getSheet("Sheet4").getRow(4).createCell(2).setCellValue("Eee");
		w.getSheet("Sheet4").getRow(4).createCell(3).setCellValue("Hosur");
		
		FileOutputStream fo=new FileOutputStream(f);
		w.write(fo);
		w.close();
	}

}
