package com.excel;

import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class Runner {

	public static void main(String[] args) {
		List<excelpojo> listBooks = new ArrayList<>();
		ExcelR e1 = new ExcelR("D:/nk0072025/TECHM/Book1.xlsx");
		try {
			listBooks = e1.readex();
			for(excelpojo s:listBooks)
			{
				System.out.println("***********************");
				System.out.println("Special printing " + s.getSub());
				System.out.println("Special printing " + s.getVar1());
	            System.out.println("Special printing " + s.getVar2());
	            System.out.println("Special printing " + s.getVar3());
	            System.out.println("Special printing " + s.getTomail());
	            System.out.println("Special printing " + s.getCcmail());
	            System.out.println("Special printing " + s.getBccmail());
	            System.out.println("Special printing " + s.getBody());
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.print( "going in errrrrrrrrrrr" );
			//e.printStackTrace();
		}
		
     	
	}

}
