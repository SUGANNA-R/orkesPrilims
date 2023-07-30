package com.handson.SpreadSheetDesign;

import com.spire.xls.*;

public class ExcelDTO {
String Value;
Workbook wk;


public ExcelDTO(Workbook wk) {
	super();
	this.wk = wk;
}
public ExcelDTO() {
	super();
	// TODO Auto-generated constructor stub
}
public String getValue() {
	return Value;
}
public void setValue(String value) {
	Value = value;
}
public Workbook getWk() {
	return wk;
}
public void setWk(Workbook wk) {
	this.wk = wk;
}

}
