package com.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelR 

{

	String excelFilePath;
	
	public ExcelR(String excelFilePath) {
		super();
		this.excelFilePath = excelFilePath;
	}
	
	private static Object getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
		return cell.getStringCellValue();
	
		case Cell.CELL_TYPE_BOOLEAN:
		return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_NUMERIC:
		return cell.getNumericCellValue();
		}
	
		return null;
		}
	
	



	public List<excelpojo> readex () throws FileNotFoundException{
	List<excelpojo> listBooks = new ArrayList<>();
	//excelpojo m1 = new excelpojo();
    FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
    
    Workbook workbook;
	try {
		workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
	
    
    Iterator<Row> iterator = firstSheet.iterator();
     
    while (iterator.hasNext()) {
        Row nextRow = iterator.next();
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        excelpojo m1 = new excelpojo();
       
        while (cellIterator.hasNext()) {
        	 Cell cell = cellIterator.next();
             int columnIndex = cell.getColumnIndex();
             switch (columnIndex) {
             case 1:
 	            System.out.println(getCellValue(cell) + "sub");
 	            m1.setSub(getCellValue(cell).toString());
 	            break;
             case 2:
             	System.out.println(getCellValue(cell)+ "var1");
             	m1.setVar1(getCellValue(cell).toString());
             	break;
             case 3:
             	System.out.println(getCellValue(cell)+ "var2");
             	m1.setVar2(getCellValue(cell).toString());
                 break;
             case 4:
             	System.out.println(getCellValue(cell)+ "var3");
             	m1.setTomail(getCellValue(cell).toString());
                 break;
             case 5:
             	System.out.println(getCellValue(cell)+ "tomail");
             	m1.setCcmail(getCellValue(cell).toString());
                 break;
             case 6:
             	System.out.println(getCellValue(cell)+ "ccmail");
             	m1.setBccmail(getCellValue(cell).toString());
                 break;
             case 7:
             	System.out.println(getCellValue(cell)+ "bcc mail");
             	m1.setBody(getCellValue(cell).toString());
                 break;
             case 8:
              	System.out.println(getCellValue(cell)+ "body");
              	m1.setBody(getCellValue(cell).toString());
                  break;
            
             }
            System.out.print(" - ");
        }
        System.out.print("  I AM ADDING NOW IN THE LIST ");
        listBooks.add(m1);

        System.out.println();
    }
     
    workbook.close();
    inputStream.close();
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
	
	return listBooks;
	
	}
	
	
	
	
}
