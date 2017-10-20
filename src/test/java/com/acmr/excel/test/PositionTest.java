package com.acmr.excel.test;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.service.ExcelService;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;

/**
 * 测试重新打开
 * 
 * @author jinhr
 *
 */
public class PositionTest {
	private ExcelFont excelFont;
	private Content content;
	private ExcelService excelService;
	private ExcelBook excelBook;
	private ReturnParam returnParam;
	private SpreadSheet spreadSheet;

	@Before
	public void before() {
		excelFont = new ExcelFont();
		content = new Content();
		excelService = new ExcelService();
		excelBook = TestUtil.createNewExcel();
		returnParam = new ReturnParam();
		spreadSheet = new SpreadSheet();
	}

	/**
	 * 测试下划线的重新打开
	 */
	@Test
	public void testPositionWithUnderline() {
		excelFont.setUnderline((byte) 0);
		content.setUnderline(excelFont.getUnderline() == 0 ? 0 : 1);
		Assert.assertEquals(Integer.valueOf(0), content.getUnderline());
		excelFont.setUnderline((byte) 1);
		content.setUnderline(excelFont.getUnderline() == 0 ? 0 : 1);
		Assert.assertEquals(Integer.valueOf(1), content.getUnderline());
		excelFont.setUnderline((byte) 2);
		content.setUnderline(excelFont.getUnderline() == 0 ? 0 : 1);
		Assert.assertEquals(Integer.valueOf(1), content.getUnderline());
	}
	/**
	 * 测试设置单元格取消锁定后的还原
	 */
	@Test
	public void testCellLockPosition(){
		ExcelCell excelCell = new ExcelCell();
		excelCell.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, excelCell);
		SpreadSheet sheet = excelService.positionExcel(excelBook.getSheets().get(0), spreadSheet, 200, returnParam,"1");
		Assert.assertEquals(Boolean.valueOf(false), sheet.getSheet().getCells().get(0).getLocked());
	}
	/**
	 * 测试设置整行取消锁定后的还原
	 */
	@Test
	public void testRowLockPosition(){
		excelBook.getSheets().get(0).getRows().get(0).getCellstyle().setLocked(false);
		SpreadSheet sheet = excelService.positionExcel(excelBook.getSheets().get(0), spreadSheet, 200, returnParam,"1");
		Assert.assertEquals(Boolean.valueOf(false), sheet.getSheet().getGlY().get(0).getOperProp().getLocked());
	}
	/**
	 * 测试设置整列取消锁定后的还原
	 */
	@Test
	public void testColLockPosition(){
		excelBook.getSheets().get(0).getCols().get(0).getCellstyle().setLocked(false);
		SpreadSheet sheet = excelService.positionExcel(excelBook.getSheets().get(0), spreadSheet, 200, returnParam,"1");
		Assert.assertEquals(Boolean.valueOf(false), sheet.getSheet().getGlX().get(0).getOperProp().getLocked());
	}

	@After
	public void after() {
		excelFont = null;
		content = null;
	}
}
