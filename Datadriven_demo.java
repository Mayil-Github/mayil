package datadriven.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.chrono.ChronoLocalDateTime;
import java.util.Date;

import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.formula.functions.NumericFunction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.DateFormatConverter;
import org.apache.poi.xssf.streaming.SheetDataWriter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Datadriven_demo {
	
	public static void main(String[] args) throws IOException {
		
		File f=new File("C:\\Users\\Admin\\eclipse-workspace\\Maven_Test\\datadriven.xlsx") ;
		FileInputStream f1=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(f1);
		Sheet sheet = w.getSheet("Sheet1");
		Row row = sheet.getRow(5);
		Cell cell = row.getCell(1);
        CellType cellType = cell.getCellType();	
		if (cellType.equals(cellType.STRING)) {
			System.out.println(cell.getStringCellValue());
			
		} else if (cellType.equals(cellType.NUMERIC)) {
			//System.out.println(cell.getNumericCellValue());
			double value = cell.getNumericCellValue();

			System.out.println(value);

			
}{

}

		
				
	
	}

}
