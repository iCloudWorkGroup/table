package com.acmr.cache;

import java.util.HashMap;
import java.util.Map;

import com.acmr.excel.model.datavalidate.Data;

public class MemoryUtil {
	private static Map<String, Data> dataValidateMap = new HashMap<String, Data>();

	public static Map<String, Data> getDataValidateMap() {
		return dataValidateMap;
	}

	public static void setDataValidateMap(Map<String, Data> dataValidateMap) {
		MemoryUtil.dataValidateMap = dataValidateMap;
	}

	public static Data getData(String excelId) {
		Data data = dataValidateMap.get(excelId);
		if (data == null) {
			data = new Data();
			dataValidateMap.put(excelId, data);
		}
		return data;
	}

}
