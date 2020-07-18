package service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import model.ProGrad;

//			Progression -1 
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	FileOutputStream out;
	int i=1;
	public HSSFWorkbook excelGenerate(ProGrad prograd, List<ProGrad> list) throws IOException {
		try {

			File myfile= new File("C:\\Users\\lenovo\\Desktop");
			
			FileInputStream file=new FileInputStream(myfile);
			
			HSSFWorkbook myworkbook= new HSSFWorkbook(file);
			
			HSSFSheet sheet = myworkbook.createSheet("ProGradDetails");
			
			Row row =sheet.createRow(0);
			
			row.createCell(0).setCellValue("ProGrad Name");
			row.createCell(1).setCellValue("ProGrad Id");
			row.createCell(2).setCellValue("ProGrad Rate");
			row.createCell(3).setCellValue("ProGrad Comment");
			row.createCell(4).setCellValue("ProGrad Recommend");
			

			for(ProGrad fillSheet: list) {
	      	 
	      	 Row nextRows = sheet.createRow(i);
	      	nextRows.createCell(0).setCellValue(fillSheet.getName());
	      	nextRows.createCell(1).setCellValue(fillSheet.getId());
	      	nextRows.createCell(2).setCellValue(fillSheet.getRate());
	      	nextRows.createCell(3).setCellValue(fillSheet.getComment());
	      	nextRows.createCell(4).setCellValue(fillSheet.getRecommend());
			
		
			}
			// Do not modify the lines given below
			 out = new FileOutputStream(myfile);
			myworkbook.write(out);
		
			return myworkbook;
			}
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
		}
		return null;
		
	}
}
