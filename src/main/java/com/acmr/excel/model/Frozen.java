package com.acmr.excel.model;

import java.io.Serializable;

public class Frozen implements Serializable {
	private int orignalCol;
	private int orignalRow;
	private int viewRow;
	private int viewCol;

	public int getOrignalCol() {
		return orignalCol;
	}

	public void setOrignalCol(int orignalCol) {
		this.orignalCol = orignalCol;
	}

	public int getOrignalRow() {
		return orignalRow;
	}

	public void setOrignalRow(int orignalRow) {
		this.orignalRow = orignalRow;
	}

	public int getViewRow() {
		return viewRow;
	}

	public void setViewRow(int viewRow) {
		this.viewRow = viewRow;
	}

	public int getViewCol() {
		return viewCol;
	}

	public void setViewCol(int viewCol) {
		this.viewCol = viewCol;
	}

}
