package com.acmr.excel.model.complete;

import java.util.List;

import com.acmr.excel.model.datavalidate.Rule;

public class ReturnData {
	private Rule rule;
	private Integer index;
	private List<String> expResult;

	public Rule getRule() {
		return rule;
	}

	public void setRule(Rule rule) {
		this.rule = rule;
	}

	public Integer getIndex() {
		return index;
	}

	public void setIndex(Integer index) {
		this.index = index;
	}

	public List<String> getExpResult() {
		return expResult;
	}

	public void setExpResult(List<String> expResult) {
		this.expResult = expResult;
	}

}
