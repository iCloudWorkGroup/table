package com.acmr.excel.model;

import java.io.Serializable;

public class Protect implements Serializable {
	private String password = "";
	private boolean isProtect;

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public boolean isProtect() {
		return isProtect;
	}

	public void setProtect(boolean isProtect) {
		this.isProtect = isProtect;
	}

}
