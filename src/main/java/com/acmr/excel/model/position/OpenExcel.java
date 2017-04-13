package com.acmr.excel.model.position;

public class OpenExcel {
	private String excelId;
	private int rowBegin;
	private int rowEnd;
	private int colBegin;
	private int colEnd;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public int getRowBegin() {
		return rowBegin;
	}

	public void setRowBegin(int rowBegin) {
		this.rowBegin = rowBegin;
	}

	public int getRowEnd() {
		return rowEnd;
	}

	public void setRowEnd(int rowEnd) {
		this.rowEnd = rowEnd;
	}

	public int getColBegin() {
		return colBegin;
	}

	public void setColBegin(int colBegin) {
		this.colBegin = colBegin;
	}

	public int getColEnd() {
		return colEnd;
	}

	public void setColEnd(int colEnd) {
		this.colEnd = colEnd;
	}

}
