package com.handson.SpreadSheetDesign;


import java.util.Scanner;

import com.spire.xls.*;

public class createSpreadSheet {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Scanner sc= new Scanner(System.in);
		System.out.print(sc);
		String flag = "Y";
		String value;
 	    String cl;
 	    ExcelGetSet xls = new ExcelGetSet();
		try {
		    //Cretae Workbook and Sheet
	        Workbook workbook = new Workbook();
	        Worksheet sheet = workbook.getWorksheets().get(0);
	        sheet.setName("InterviewHandsOnSheet");
	        	
	        //Read input for set cell values
	        do {
	        
	        		ExcelDTO obj = new ExcelDTO();
	        		System.out.println("\nEnter Cell Range : ");
	        		cl = sc.next();
	        		System.out.println("Enter Cell Value (for formula please prefix with '=' symbol: )");
	        		value = sc.next();
	        		obj.setValue(value);
	        		obj.setWk(workbook);
	        		xls.setCellValue(cl, obj);
	        		System.out.println("Enter (Y/N) - Y to continue / N to exit : ");
	        		flag = sc.next();
	        		
	        }while(flag.equalsIgnoreCase("Y"));
	        

    		
	        //Save the resultant file
	        workbook.saveToFile("..//..//Documents/OrkesPrilims.xlsx", FileFormat.Version2013);
	        System.out.println("\n*****Spreadsheet created with the given inputs please check file in Documents Folder*****\n");
		      
	        flag = "Y";
	        do {
		        
        		System.out.println("Enter Cell Range to get its value: ");
        		cl = sc.next();
        		
        		int res = xls.getCellValue(cl);
        		 System.out.println("The value of cell " +cl+" = "+ res+"\n");
        		System.out.println("Enter (Y/N) - Y to continue / N to exit : ");
        		flag = sc.next();
        		
        }while(flag.equalsIgnoreCase("Y"));
			
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("\n*****Please execute application again and enter valid inputs*****\n");
		}
		
		

	}

}
