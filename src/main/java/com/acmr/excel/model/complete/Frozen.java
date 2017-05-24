package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Frozen implements Serializable{
	private String state;
	private Integer row;
	private Integer col;
	private String displayAreaStartAlaisX;
	private String displayAreaStartAlaisY;

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}
	

	public Integer getRow() {
		return row;
	}

	public void setRow(Integer row) {
		this.row = row;
	}

	public Integer getCol() {
		return col;
	}

	public void setCol(Integer col) {
		this.col = col;
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
