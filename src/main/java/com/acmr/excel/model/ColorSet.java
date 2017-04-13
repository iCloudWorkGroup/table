package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class ColorSet implements Serializable{
	private int sheetId;
	private List<Coordinate> coordinates;
	private String bgcolor;

	public int getSheetId() {
		return sheetId;
	}

	public void setSheetId(int sheetId) {
		this.sheetId = sheetId;
	}

	public List<Coordinate> getCoordinates() {
		return coordinates;
	}

	public void setCoordinates(List<Coordinate> coordinates) {
		this.coordinates = coordinates;
	}

	public String getBgcolor() {
		return bgcolor;
	}

	public void setBgcolor(String bgcolor) {
		this.bgcolor = bgcolor;
	}

}
