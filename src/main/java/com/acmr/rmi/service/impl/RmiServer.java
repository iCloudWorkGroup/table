package com.acmr.rmi.service.impl;

import java.rmi.RemoteException;
import java.rmi.registry.LocateRegistry;
import java.rmi.registry.Registry;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;
import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;

import com.acmr.rmi.service.RmiService;

public class RmiServer implements ServletContextListener {
//	public static void main(String[] args) {
//
//		try {
//			RmiService rmi = new RmiServiceImpl();
//			LocateRegistry.createRegistry(10999);
//			try {
//				Naming.bind("rmi://127.0.0.1:10999/RmiService", rmi);
//				System.out.println(">>>>>INFO:远程IHello对象绑定成功！");
//			} catch (MalformedURLException e) {
//				e.printStackTrace();
//			} catch (AlreadyBoundException e) {
//				e.printStackTrace();
//			}
//		} catch (RemoteException e) {
//			e.printStackTrace();
//		}
//	}

	@Override
	public void contextDestroyed(ServletContextEvent arg0) {

	}

	@Override
	public void contextInitialized(ServletContextEvent arg0) {
		try {
			RmiService rmi = new RmiServiceImpl();
			LocateRegistry.createRegistry(11999);
			Context namingContext = new InitialContext();
			namingContext.rebind("rmi://127.0.0.1:11999/RmiService", rmi);
			System.out.println("rmi server start");
		} catch (RemoteException e) {
			e.printStackTrace();
		} catch (NamingException e) {
			e.printStackTrace();
		}

	}
}
