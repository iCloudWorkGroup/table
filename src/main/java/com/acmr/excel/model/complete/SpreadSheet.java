package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.AreaSet;

import acmr.excel.pojo.ExcelDataValidation;

public class SpreadSheet implements Serializable {

	private SheetElement sheet = new SheetElement();
	private String name;
	private int sort;
	private String tempHTML;
	private List<AreaSet> validate = new ArrayList<AreaSet>();

	public SheetElement getSheet() {
		return sheet;
	}

	public void setSheet(SheetElement sheet) {
		this.sheet = sheet;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public String getTempHTML() {
		return tempHTML;
	}

	public void setTempHTML(String tempHTML) {
		this.tempHTML = tempHTML;
	}

	public List<AreaSet> getValidate() {
		return validate;
	}

	public void setValidate(List<AreaSet> validate) {
		this.validate = validate;
	}

}
