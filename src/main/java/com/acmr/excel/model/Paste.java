package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class Paste implements Serializable {
	private String excelId;
	private int startColSort;
	private int startRowSort;
	private List<OuterPasteData> pasteData;
	private int colLen;
	private int rowLen;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public List<OuterPasteData> getPasteData() {
		return pasteData;
	}

	public void setPasteData(List<OuterPasteData> pasteData) {
		this.pasteData = pasteData;
	}

	public int getColLen() {
		return colLen;
	}

	public void setColLen(int colLen) {
		this.colLen = colLen;
	}

	public int getRowLen() {
		return rowLen;
	}

	public void setRowLen(int rowLen) {
		this.rowLen = rowLen;
	}

	public int getStartColSort() {
		return startColSort;
	}

	public void setStartColSort(int startColSort) {
		this.startColSort = startColSort;
	}

	public int getStartRowSort() {
		return startRowSort;
	}

	public void setStartRowSort(int startRowSort) {
		this.startRowSort = startRowSort;
	}

}
