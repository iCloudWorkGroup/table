package com.acmr.excel.service.impl;

import org.springframework.stereotype.Service;

import com.acmr.excel.model.Constant;
import com.acmr.excel.service.StoreService;

//@Service
public class MapServiceImpl implements StoreService{

	@Override
	public Object get(String id) {
		return Constant.map.get(id);
	}

	@Override
	public Object set(String id, Object object) {
		Constant.map.put(id, object);
		return null;
	}

}
