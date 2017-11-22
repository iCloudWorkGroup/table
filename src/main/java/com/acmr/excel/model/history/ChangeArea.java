package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class ChangeArea implements Serializable {
	private int colIndex;
	private int rowIndex;
	private List<Object> originalValues = new ArrayList<Object>();
	private List<Object> updateValues = new ArrayList<Object>();
	private boolean isExist = true;
	private int type = 0;

	public int getType() {
		return type;
	}

	public void setType(int type) {
		this.type = type;
	}

	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public List<Object> getOriginalValues() {
		return originalValues;
	}

	public void setOriginalValues(List<Object> originalValues) {
		this.originalValues = originalValues;
	}

	public List<Object> getUpdateValues() {
		return updateValues;
	}

	public void setUpdateValues(List<Object> updateValues) {
		this.updateValues = updateValues;
	}

	public boolean isExist() {
		return isExist;
	}

	public void setExist(boolean isExist) {
		this.isExist = isExist;
	}

	public ChangeArea() {
		this.originalValues.add(null);
		this.originalValues.add(null);
		this.updateValues.add(null);
		this.updateValues.add(null);
	}
}
