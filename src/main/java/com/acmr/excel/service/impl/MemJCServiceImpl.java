package com.acmr.excel.service.impl;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.service.StoreService;
import com.danga.MemCached.MemCachedClient;

@Service
public class MemJCServiceImpl implements StoreService {
	@Resource
	private MemCachedClient memCachedClient;

	@Override
	public Object get(String id) {
		return memCachedClient.get(id);
	}

	@Override
	public Object set(String id, Object object) {
		memCachedClient.set(id, object);
		return null;
	}

}
