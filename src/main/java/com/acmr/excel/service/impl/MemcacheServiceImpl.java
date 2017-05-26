package com.acmr.excel.service.impl;

import javax.annotation.Resource;

import net.spy.memcached.MemcachedClient;
import net.spy.memcached.internal.OperationFuture;

import org.springframework.stereotype.Service;

import com.acmr.excel.model.Constant;
import com.acmr.excel.service.StoreService;

//@Service
public class MemcacheServiceImpl implements StoreService {
	@Resource
	private MemcachedClient memcachedClient;
	private int exp = Integer.valueOf(Constant.MEMCACHED_EXP_TIME);

	@Override
	public Object get(String id) {
		return memcachedClient.get(id);
	}

	@Override
	public Object set(String id, Object object) {
		OperationFuture<Boolean> f = memcachedClient.set(id, exp, object);
		return null;
	}

}
