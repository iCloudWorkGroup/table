package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

import com.acmr.excel.model.datavalidate.Rule;

public class AreaSet implements Serializable {
	private List<Coordinate> coordinate;
	private String color;
	private boolean lock;
	private Rule rule;

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

	public boolean isLock() {
		return lock;
	}

	public void setLock(boolean lock) {
		this.lock = lock;
	}

	public Rule getRule() {
		return rule;
	}

	public void setRule(Rule rule) {
		this.rule = rule;
	}

}
