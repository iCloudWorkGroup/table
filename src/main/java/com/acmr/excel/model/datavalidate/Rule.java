package com.acmr.excel.model.datavalidate;

import java.io.Serializable;

public class Rule implements Serializable {
	private int validationType;
	private String formula1;
	private String formula2;
	private Integer index;

	public int getValidationType() {
		return validationType;
	}

	public void setValidationType(int validationType) {
		this.validationType = validationType;
	}

	public String getFormula1() {
		return formula1;
	}

	public void setFormula1(String formula1) {
		this.formula1 = formula1;
	}

	public String getFormula2() {
		return formula2;
	}

	public void setFormula2(String formula2) {
		this.formula2 = formula2;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result
				+ ((formula1 == null) ? 0 : formula1.hashCode());
		result = prime * result
				+ ((formula2 == null) ? 0 : formula2.hashCode());
		result = prime * result + validationType;
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Rule other = (Rule) obj;
		if (formula1 == null) {
			if (other.formula1 != null)
				return false;
		} else if (!formula1.equals(other.formula1))
			return false;
		if (formula2 == null) {
			if (other.formula2 != null)
				return false;
		} else if (!formula2.equals(other.formula2))
			return false;
		if (validationType != other.validationType)
			return false;
		return true;
	}

	public Integer getIndex() {
		return index;
	}

	public void setIndex(Integer index) {
		this.index = index;
	}
	
	
}
