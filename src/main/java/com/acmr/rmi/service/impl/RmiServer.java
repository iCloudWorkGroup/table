package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;
import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import com.acmr.excel.util.PropertiesReaderUtil;
import com.acmr.rmi.service.RmiService;

public class RmiServer implements ServletContextListener {

	@Override
	public void contextDestroyed(ServletContextEvent arg0) {

	}

	@Override
	public void contextInitialized(ServletContextEvent arg0) {
		try {
			RmiService rmi = new RmiServiceImpl();
			String port = PropertiesReaderUtil.get("rmi.port");
			LocateRegistry.createRegistry(Integer.valueOf(port));
			Context namingContext = new InitialContext();
			namingContext.rebind("rmi://127.0.0.1:" + port + "/RmiService", rmi);
			System.out.println("rmi server start");
		} catch (RemoteException e) {
			e.printStackTrace();
		} catch (NamingException e) {
			e.printStackTrace();
		}

	}
}
