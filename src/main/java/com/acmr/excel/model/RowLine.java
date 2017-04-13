package com.acmr.excel.model;

import java.io.Serializable;

public class RowLine implements Serializable{
	private String excelId;
	private String rowNum;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public String getRowNum() {
		return rowNum;
	}

	public void setRowNum(String rowNum) {
		this.rowNum = rowNum;
	}
	
}
