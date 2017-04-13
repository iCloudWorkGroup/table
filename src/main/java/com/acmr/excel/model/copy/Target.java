package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Target implements Serializable {
	private int orignalCol;
	private int orignalRow;

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

}
