package com.cg.hims.scor.export;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.cg.hims.scor.entity.StudentInfo;

public class xlsxFileCreater {

	public static ByteArrayInputStream contactListToExcelFile(List<StudentInfo> studentinfo)
	{   
	try{
			
			
	   final String[] columns = {"studid","firstName","lastName","Email"};
		
		Workbook workbook=new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("relo");
		Font headerFont=workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 17);
		headerFont.setColor(IndexedColors.BLUE.getIndex());
		headerFont.setFontName("Arial");
		
		
		Font rowfont=workbook.createFont();
		rowfont.setBold(false);
		rowfont.setFontHeightInPoints((short) 12);
		rowfont.setColor(IndexedColors.ORANGE.getIndex());
		rowfont.setFontName("Arial");
		
/************************************************************************************************/			
		CellStyle headeCellStyle= workbook.createCellStyle();
		headeCellStyle.setFont(headerFont);
		headeCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		headeCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		
		
	    CellStyle rowStyle= workbook.createCellStyle();
	    rowStyle.setFont(rowfont);
	    rowStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
	    rowStyle.setAlignment(CellStyle.ALIGN_CENTER);
//	    rowStyle.setAlignment(HorizontalAlignment.CENTER);
//	    rowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
		Row headerRow=sheet.createRow(0);
			for(int i=0;i<columns.length;i++)
			{
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columns[i]);
				cell.setCellStyle(headeCellStyle);
							
			}
			int rNum = 1;
			for(StudentInfo str : studentinfo)
			{
				Row row = sheet.createRow(rNum++);
		    	row.createCell(0).setCellValue(str.getStudid());
		    	row.createCell(1).setCellValue(str.getFirstName());
		    	row.createCell(2).setCellValue(str.getLastName());
		    	row.createCell(3).setCellValue(str.getEmail());
		    
			}	

			for(int i=0;i<columns.length;i++)
		    {
		    	
		    	sheet.autoSizeColumn(i);
		    }
			 ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		        workbook.write(outputStream);
		        return new ByteArrayInputStream(outputStream.toByteArray());
			} catch (IOException ex) {
				ex.printStackTrace();
				return null;
			}
	
		
	}
	
	
	
}
