package com.acmr.excel.model.datavalidate;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class Data {
	private Map<String, Integer> rowMap = new HashMap<String, Integer>();
	private Map<String, Integer> colMap = new HashMap<String, Integer>();
	private Map<String, Map<String, Integer>> cellMap = new HashMap<String, Map<String, Integer>>();
	private List<Rule> ruleList = new ArrayList<Rule>();


	public List<Rule> getRuleList() {
		return ruleList;
	}

	public void setRuleList(List<Rule> ruleList) {
		this.ruleList = ruleList;
	}

	public Map<String, Integer> getRowMap() {
		return rowMap;
	}

	public void setRowMap(Map<String, Integer> rowMap) {
		this.rowMap = rowMap;
	}

	public Map<String, Integer> getColMap() {
		return colMap;
	}

	public void setColMap(Map<String, Integer> colMap) {
		this.colMap = colMap;
	}

	public Map<String, Map<String, Integer>> getCellMap() {
		return cellMap;
	}

	public void setCellMap(Map<String, Map<String, Integer>> cellMap) {
		this.cellMap = cellMap;
	}
	
}
