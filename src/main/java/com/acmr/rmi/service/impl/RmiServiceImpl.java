package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.server.UnicastRemoteObject;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

import javax.annotation.Resource;
import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;

import net.spy.memcached.MemcachedClient;
import net.spy.memcached.internal.OperationFuture;

import org.springframework.stereotype.Service;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelDataValidation;

import com.acmr.cache.MemoryUtil;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.util.DataValidateUtil;
import com.acmr.rmi.service.RmiService;

public class RmiServiceImpl extends UnicastRemoteObject implements RmiService {
	private MemcachedClient memCachedClient;
	
	public RmiServiceImpl(MemcachedClient memCachedClient) throws RemoteException {
		this.memCachedClient = memCachedClient;
	} 
	
	

	@Override
	public boolean saveExcelBook(String excelId, ExcelBook excelBook) throws RemoteException {
		//MemcachedClient memcachedClient = MemcacheFactory.CACHESOURCE.getMemcacheClient();
		OperationFuture<Boolean> excelResult = memCachedClient.set(excelId, 60 * 60 * 24 * 1, excelBook);
		OperationFuture<Boolean> opeResult = memCachedClient.set(excelId+"_ope", 60 * 60 * 24 * 1, 0);
		try {
			if(excelResult.get() && opeResult.get()){
				return true;
			}
		} catch (InterruptedException e) {
			e.printStackTrace();
		} catch (ExecutionException e) {
			e.printStackTrace();
		}
		return false;
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
		List<ExcelDataValidation> excelDataValidations = new ArrayList<ExcelDataValidation>();
		Data data = MemoryUtil.getDataValidateMap().get(excelId);
		if(data != null){
			DataValidateUtil.map2List(data, excelDataValidations, excelBook.getSheets().get(0));
			if(excelDataValidations.size() >0){
				excelBook.getSheets().get(0).setExcelDataValidations(excelDataValidations);
			}
		}
		return excelBook;
	}
	

}
