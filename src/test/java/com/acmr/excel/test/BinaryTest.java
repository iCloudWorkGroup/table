package com.acmr.excel.test;
import java.util.ArrayList;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.util.BinarySearch;
import com.alibaba.fastjson.JSON;

/**
 * 二分查询测试
 * @author jinhr
 *
 */
public class BinaryTest {
	
	/**
	 * 定义初始化测试数组
	 */
	private List<Gly> glyList = new ArrayList<Gly>();
	
	/**
	 * 在测试方法执行前执行
	 * @throws Exception
	 */
	
	@Before
	public void setUp() throws Exception {
		for (int i = 0; i < 100; i++) {
			Gly gly = new Gly();
			gly.setAliasY(i + 1 + "");
			gly.setHeight(19);
			gly.setTop(i * 19 + i);
			glyList.add(gly);
		}
		//System.out.println(JSON.toJSONString(glyList));
	}
	
	@Test
	public void testBinary() {
		findNoexistInRange();
		findexistNoInRange();
		findExcessMaxTop();
	}
	
	/**
	 * 测试在范围内不存在的top的情况
	 */
	private void findNoexistInRange(){
		int rowIndex = BinarySearch.rowsBinarySearch(glyList, 101);
		System.out.println(rowIndex);
	}
	/**
	 * 测试在范围内存在的top的情况
	 */
	private void findexistNoInRange(){
		int rowIndex = BinarySearch.rowsBinarySearch(glyList, 100);
		System.out.println(rowIndex);
	}
	/**
	 * 测试超出最大top的情况
	 */
	private void findExcessMaxTop (){
		int rowIndex = BinarySearch.rowsBinarySearch(glyList, 1985);
		System.out.println(rowIndex);
	}
}
