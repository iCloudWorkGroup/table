package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.server.UnicastRemoteObject;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.cache.MemcacheFactory;
import com.acmr.excel.model.Constant;
import com.acmr.rmi.service.RmiService;

public class RmiServiceImpl extends UnicastRemoteObject implements RmiService {
	
	
	public RmiServiceImpl() throws RemoteException {
	} 
	
	@Override
	public void saveExcelBook(String excelId, ExcelBook excelBook) throws RemoteException {
		MemcachedClient memcachedClient = MemcacheFactory.CACHESOURCE.getMemcacheClient();
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME,excelBook);
		memcachedClient.set(excelId+"_ope", Constant.MEMCACHED_EXP_TIME, 0);
		System.out.println(memcachedClient.get(excelId+"_ope"));
	}

	@Override
	public ExcelBook getExcelBook(String excelId, int step) throws RemoteException {
		MemcachedClient	memcachedClient = MemcacheFactory.CACHESOURCE.getMemcacheClient();
		ExcelBook excelBook = null;
		int curStep = (int) memcachedClient.get(excelId + "_ope");
		if (curStep == step) {
			excelBook = (ExcelBook) memcachedClient.get(excelId);
		} else {
			for (int i = 0; i < 100; i++) {
				int st = (int) memcachedClient.get(excelId + "_ope");
				if (step == st) {
					excelBook = (ExcelBook) memcachedClient.get(excelId);
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
	public static void main(String[] args) {
		try {
			RmiServiceImpl r = new RmiServiceImpl();
			r.saveExcelBook("1", new ExcelBook());
		} catch (RemoteException e) {
			e.printStackTrace();
		}
	}
	

}
