package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.server.UnicastRemoteObject;

import javax.annotation.Resource;
import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;

import org.springframework.stereotype.Service;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.cache.MemcacheFactory;
import com.acmr.excel.model.Constant;
import com.acmr.excel.service.StoreService;
import com.acmr.rmi.service.RmiService;
import com.danga.MemCached.MemCachedClient;

public class RmiServiceImpl extends UnicastRemoteObject implements RmiService {
	private MemCachedClient memCachedClient;
	
	public RmiServiceImpl(MemCachedClient memCachedClient) throws RemoteException {
		this.memCachedClient = memCachedClient;
	} 
	
	

	@Override
	public void saveExcelBook(String excelId, ExcelBook excelBook) throws RemoteException {
		//MemcachedClient memcachedClient = MemcacheFactory.CACHESOURCE.getMemcacheClient();
		memCachedClient.set(excelId,excelBook);
		memCachedClient.set(excelId+"_ope",  0);
		//System.out.println(storeService.get(excelId+"_ope"));
	}

	@Override
	public ExcelBook getExcelBook(String excelId, int step) throws RemoteException {
		//MemcachedClient	memcachedClient = MemcacheFactory.CACHESOURCE.getMemcacheClient();
		ExcelBook excelBook = null;
		int curStep = (int) memCachedClient.get(excelId + "_ope");
		if (curStep == step) {
			excelBook = (ExcelBook) memCachedClient.get(excelId);
		} else {
			for (int i = 0; i < 100; i++) {
				int st = (int) memCachedClient.get(excelId + "_ope");
				if (step == st) {
					excelBook = (ExcelBook) memCachedClient.get(excelId);
				} else {
					try {
						Thread.sleep(100);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
				}
			}
		}
		return excelBook;
	}
	

}
