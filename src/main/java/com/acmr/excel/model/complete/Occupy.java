package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Occupy implements Serializable {
	private List<String> x = new ArrayList<String>();
	private List<String> y = new ArrayList<String>();

	public List<String> getX() {
		return x;
	}

	public void setX(List<String> x) {
		this.x = x;
	}

	public List<String> getY() {
		return y;
	}

	public void setY(List<String> y) {
		this.y = y;
	}
}
