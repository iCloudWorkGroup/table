package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Frozen implements Serializable{
	private String state;
	private String rowIndex;
	private String colIndex;
	private String displayAreaStartAlaisX;
	private String displayAreaStartAlaisY;

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	public String getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(String rowIndex) {
		this.rowIndex = rowIndex;
	}

	public String getColIndex() {
		return colIndex;
	}

	public void setColIndex(String colIndex) {
		this.colIndex = colIndex;
	}

	public String getDisplayAreaStartAlaisX() {
		return displayAreaStartAlaisX;
	}

	public void setDisplayAreaStartAlaisX(String displayAreaStartAlaisX) {
		this.displayAreaStartAlaisX = displayAreaStartAlaisX;
	}

	public String getDisplayAreaStartAlaisY() {
		return displayAreaStartAlaisY;
	}

	public void setDisplayAreaStartAlaisY(String displayAreaStartAlaisY) {
		this.displayAreaStartAlaisY = displayAreaStartAlaisY;
	}

}
