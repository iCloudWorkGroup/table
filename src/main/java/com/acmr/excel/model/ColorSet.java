package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class ColorSet implements Serializable {
	private List<Coordinate> coordinate;
	private String color;

	public List<Coordinate> getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(List<Coordinate> coordinate) {
		this.coordinate = coordinate;
	}

	public String getColor() {
		return color;
	}

	public void setColor(String color) {
		this.color = color;
	}

}
