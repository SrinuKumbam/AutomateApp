package com.example.Automate.businessLogic;

/*import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import com.relevantcodes.extentreports.NetworkMode;*/
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.util.logging.Logger;

@Service
public class ExcelSheetService {
	
	private String resultName = "";
	private String result = "";
//	private FileInputStream fileInputStream1;
	private HSSFWorkbook workbook1;
	
	public void compareTwoFiles(MultipartFile file1, MultipartFile file2) throws IOException {
		
		try {
			//fileInputStream1 = file1;
	        workbook1 = new HSSFWorkbook(file1.getInputStream());
	        HSSFSheet worksheet1 = workbook1.getSheet("Sheet2");

	        int rowCount1= worksheet1.getPhysicalNumberOfRows();    

	        HSSFWorkbook workbook2 = new HSSFWorkbook(file2.getInputStream());
	        HSSFSheet worksheet2 = workbook2.getSheet("Sheet2");

	        int rowCount2= worksheet2.getPhysicalNumberOfRows();

			if(rowCount1 >= 0) {
			    HSSFRow row1 = worksheet1.getRow(1);
			    
			    String idstr1 = "";
			    HSSFCell id1 = row1.getCell(1);
			    if (id1 != null) {
			        id1.setCellType(CellType.STRING);
			        idstr1 = id1.getStringCellValue();
			    }
			    resultName = idstr1;
			    
			    HSSFRow cityRow = worksheet1.getRow(4);
			    HSSFCell city1 = cityRow.getCell(1);
			    String cityName1 = "";
			    String cityName2 = "";
			    
			    if (city1 != null) {
			        city1.setCellType(CellType.STRING);
			        cityName1 = city1.getStringCellValue();
			    }
			    System.out.println(idstr1);
			    
			    for (int i = 1; i < rowCount2; i++) {
			        HSSFRow row2 = worksheet2.getRow(i);
			        
			        String idstr2 = "";
			        HSSFCell id2 = row2.getCell(1);
			        if (id2 != null) {
			            id2.setCellType(CellType.STRING);
			            idstr2 = id2.getStringCellValue();
			        }
			        		
			        if(!idstr1.equals(idstr2))
			        {
			            continue;
			        }
			        else
			        {
			        	HSSFCell city2 = row2.getCell(4);
				        if (city2 != null) {
				        	city2.setCellType(CellType.STRING);
				            cityName2 = city2.getStringCellValue();
				        	System.out.println("City name is "+cityName2);
				        }
			        }
			        
			        if(cityName1.equalsIgnoreCase(cityName2)) {
			        	result = "City name Matching";
			        	System.out.println("City names from book1 "+cityName1+" City name from book2 "+cityName2);
			        }else
			        {
			        	result = "City name is not matching";
			        	System.out.println("City names from book1 "+cityName1+" City name from book2 "+cityName2);
			        }
			        
			        createResultFile(cityName1, cityName2);
			    }
			}
		}catch(Exception e) {
			e.printStackTrace();
		}
		finally {
			workbook1.close();
		}
	}
	
	private FileOutputStream fileOut;
	public void createResultFile(String city1, String city2) throws IOException {
		try {
        Workbook resultBook = new HSSFWorkbook();
        Sheet resultSheet = resultBook.createSheet("First");
        
		Font headerFont = resultBook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short)12);
		headerFont.setColor(IndexedColors.BLACK.index);
		//Create a CellStyle with the font
		CellStyle headerStyle = resultBook.createCellStyle();
		headerStyle.setFont(headerFont);
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
    
		Row headerRow = resultSheet.createRow(0);
		String headers[] = {"Name", "Book1 city name", "Book2 city name", "Status"};
			
		for(int i=0; i< headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(headerStyle);
		}
			
		Row row = resultSheet.createRow(1);
			row.createCell(0).setCellValue(resultName);
			row.createCell(1).setCellValue(city1);
			row.createCell(2).setCellValue(city2);
			row.createCell(3).setCellValue(result);
			
		fileOut = new FileOutputStream("C:/Practice/program/Result.xls");
			resultBook.write(fileOut);
			fileOut.close();
			resultBook.close();
			System.out.println("Completed");	
		}catch(Exception e) {
			e.printStackTrace();
		}
		finally {
			fileOut.close();
		}
	}
}