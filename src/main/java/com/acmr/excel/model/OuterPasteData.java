package com.acmr.excel.model;

import java.io.Serializable;

public class OuterPasteData implements Serializable {
	private int colSort;
	private int rowSort;
	private String text;

	public int getColSort() {
		return colSort;
	}

	public void setColSort(int colSort) {
		this.colSort = colSort;
	}

	public int getRowSort() {
		return rowSort;
	}

	public void setRowSort(int rowSort) {
		this.rowSort = rowSort;
	}

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}
}
