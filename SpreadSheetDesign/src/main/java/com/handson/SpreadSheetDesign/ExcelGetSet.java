package com.handson.SpreadSheetDesign;

import com.spire.xls.*;

public class ExcelGetSet 
{
	

	public int getCellValue(String cellId) {
		// TODO Auto-generated method stub
		try {
		//Create a Workbook instance
        Workbook workbook = new Workbook();
        int res;
        //Load an Excel sample document
        workbook.loadFromFile( "..//..//Documents/OrkesPrilims.xlsx");

		 Worksheet sheet = workbook.getWorksheets().get(0);
		 CellRange cell = sheet.getRange().get(cellId);
		 if(cell.hasFormula()) {
			 res = (int)cell.getFormulaNumberValue();
		 }
		 else {
			 res = Integer.parseInt(cell.getValue());
		 }
   	
		return res;
		}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("\n*****Please open the Spreadsheet and look for the CellValue*****\n");
			return 0;
		}
	}
	public void setCellValue(String cellId, ExcelDTO obj) {
		// TODO Auto-generated method stub
		try {
			
		 Worksheet sheet = obj.getWk().getWorksheets().get(0);
		 int res;
		 CellRange cell = sheet.getRange().get(cellId);
         cell.setText(obj.getValue());
         if(cell.hasFormula()) {
			 res = (int)cell.getFormulaNumberValue();
		 }
		 else {
			 res = Integer.parseInt(cell.getValue());
		 }
         System.out.println("Cell value stored "+cellId+ " = "+res+"\n\n");

	      
	       
	}
		catch(Exception e)
		{
			e.printStackTrace();
			System.out.println("\n*****Invalid input*****\n");
		}
	}
}
