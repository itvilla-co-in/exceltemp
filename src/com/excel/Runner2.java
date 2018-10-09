package com.excel;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Runner2 {

	public static void main(String[] args) {
		
		ExcelPOIHelper e = new ExcelPOIHelper();
		String location = "D:/nk0072025/TECHM/images/emp.xlsx";
		Map<Integer, List<MyCell>> excelmap = new HashMap<>();;
     	try {
     		
     	excelmap = e.readExcel(location);
     	
        // using for-each loop for iteration over Map.entrySet() 
        for (Map.Entry<Integer, List<MyCell>> entry : excelmap.entrySet())
        {
            System.out.println("Key = " + entry.getKey() + 
                             ", Value = " + entry.getValue()); 
            List<MyCell> temp = new ArrayList<>();
            temp = entry.getValue();
            
            for(MyCell mc: temp) {
                System.out.print(mc.getContent());
                System.out.println("   ");
            }
            System.out.println();
     	
        }
     	
		} catch (IOException e1) {
			System.out.println("Error" + e1.toString()); 
			e1.printStackTrace();
		}
	}

}
