package com.acmr.excel.util;

import com.acmr.excel.model.Constant;

import net.spy.memcached.MemcachedClient;

public class MemcacheUtil {
	
   /**
    * 从缓存中取对象
    * @param key  键
    * @param memcache MemcachedClient
    * @return 对象实体
    * @throws InterruptedException
    */
	
   public static synchronized Object get(String key,MemcachedClient memcache) throws InterruptedException{
		String noLock = (String) memcache.get(key + "PK");
		if ("true".equals(noLock)) {
			memcache.set(key + "PK", Constant.MEMCACHED_EXP_TIME, "false");
			return memcache.get(key);
		} else {
			for (int i = 0; i < 1000; i++) {
				noLock = (String) memcache.get(key + "PK");
				if ("true".equals(noLock)) {
					memcache.set(key + "PK", Constant.MEMCACHED_EXP_TIME, "false");
					return memcache.get(key);
				} else {
					Thread.sleep(10);
				}
			}
			return memcache.get(key);
		}
   }
   /**
    * 存对象到缓存中
    * @param hey 键
    * @param memcache MemcachedClient
    * @param o 对象
    */
   
	public static void set(String hey, MemcachedClient memcache, Object object) {
		memcache.set(hey, Constant.MEMCACHED_EXP_TIME, object);
		memcache.set(hey + "PK", Constant.MEMCACHED_EXP_TIME, "true");
	}
}
