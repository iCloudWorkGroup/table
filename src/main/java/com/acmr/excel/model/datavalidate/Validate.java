package com.acmr.excel.model.datavalidate;

import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.AreaSet;

public class Validate {
	private List<AreaSet> rules = new ArrayList<AreaSet>();
	private int total;

	public int getTotal() {
		return total;
	}

	public void setTotal(int total) {
		this.total = total;
	}

	public List<AreaSet> getRules() {
		return rules;
	}

	public void setRules(List<AreaSet> rules) {
		this.rules = rules;
	}

}
