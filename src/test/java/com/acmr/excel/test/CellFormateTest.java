package com.acmr.excel.test;


import java.util.Date;

import org.junit.Assert;
import org.junit.Test;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelCell;

import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.util.CellFormateUtil;


public class CellFormateTest {
	/**
	 * 测试是否为数字
	 */
	@Test
	public void testIsNumeric() {
//		Assert.assertEquals(true, CellFormateUtil.isNumeric("1.1"));
//		Assert.assertEquals(true, CellFormateUtil.isNumeric("1"));
//		Assert.assertEquals(true, CellFormateUtil.isNumeric("123"));
//		Assert.assertEquals(true, CellFormateUtil.isNumeric("123.23"));
//		Assert.assertEquals(false, CellFormateUtil.isNumeric("a123"));
//		Assert.assertEquals(false, CellFormateUtil.isNumeric("&123"));
//		Assert.assertEquals(false, CellFormateUtil.isNumeric("123."));
//		Assert.assertEquals(false, CellFormateUtil.isNumeric("123.23.02"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("123"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("123.123"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("+123.123"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("-123.123"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("123.0120"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("0.0120"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric("0.3"));
		Assert.assertEquals(false, CellFormateUtil.isNumeric("123.12.3"));
		Assert.assertEquals(true, CellFormateUtil.isNumeric(".3"));
		Assert.assertEquals(false, CellFormateUtil.isNumeric("@123"));
		Assert.assertEquals(false, CellFormateUtil.isNumeric("AB"));
		Assert.assertEquals(false, CellFormateUtil.isNumeric("AB,AB"));
	}
	/**
	 * 测试设置小数点
	 */
	@Test
	public void testSetDecimalPoint() {
		//整数，有小数点且位数正常
		String dp1 = CellFormateUtil.setDecimalPoint(2, "123");
		Assert.assertEquals("123.00",dp1);
		//整数，无小数点
		String dp2 = CellFormateUtil.setDecimalPoint(0, "123");
		Assert.assertEquals("123",dp2);
		//整数，小数点位数不正常
		String dp3 = CellFormateUtil.setDecimalPoint(-1, "123");
		Assert.assertEquals(null,dp3);
		//小数，有小数点且位数正常
		String dp4 = CellFormateUtil.setDecimalPoint(2, "123.023");
		Assert.assertEquals("123.02", dp4);
		//小数，无小数点
		String dp5 = CellFormateUtil.setDecimalPoint(0, "123.023");
		Assert.assertEquals("123", dp5);
		// 整数，小数点位数不正常
		String dp6 = CellFormateUtil.setDecimalPoint(-1, "123.023");
		Assert.assertEquals(null, dp6);
		
	}
	/**
	 * 测试千分位
	 */
	@Test
	public void testSetThousandPoint() {
		//整数有千分位
		Assert.assertEquals("1,231,231", CellFormateUtil.setThousandPoint(true, "1231231"));
		//整数无千分位
		Assert.assertEquals("1231231", CellFormateUtil.setThousandPoint(false, "1231231"));
		//小数有千分位
		Assert.assertEquals("1,231,231.23", CellFormateUtil.setThousandPoint(true, "1231231.23"));
		//小数无千分位
		Assert.assertEquals("1231231.23", CellFormateUtil.setThousandPoint(false, "1231231.23"));
	}
	/**
	 * 测试日期
	 */
	@Test
	public void testIsValidDate() {
		//Calendar.getInstance().set(2015, 04, 27); 
		Assert.assertNull(CellFormateUtil.getDate("2015-04-32", "yyyy-MM-dd"));
		Assert.assertNull(CellFormateUtil.getDate("2015-02-29", "yyyy-MM-dd"));
		Assert.assertNotNull(CellFormateUtil.getDate("2012-02-29", "yyyy-MM-dd"));
		System.out.println(CellFormateUtil.getDate("2012-02-29", "yyyy-MM-dd"));
//		Assert.assertEquals(true, CellFormateUtil.getDate("2015-04", "yyyy-mm"));
//		Assert.assertEquals(true, CellFormateUtil.getDate("2015", "yyyy"));
//		Assert.assertEquals(true, CellFormateUtil.getDate("2015年04月27日", "yyyy年mm月dd日"));
//		Assert.assertEquals(true, CellFormateUtil.getDate("2015年04月", "yyyy年mm月"));
//		Assert.assertEquals(true, CellFormateUtil.getDate("2015年", "yyyy年"));
	}
	/**
	 * 测试数字
	 */
	@Test
	public void testSetNumber(){
		//小数设置4位小数点和千分位
		ExcelCell excelCell = new ExcelCell();
		excelCell.setText("1231231.23");
		CellFormateUtil.setNumber(excelCell, 4, true);
		Assert.assertEquals("1,231,231.2300", excelCell.getShowText());
		//整数设置4位小数点和千分位
		excelCell.setText("23232323");
		CellFormateUtil.setNumber(excelCell, 4, true);
		Assert.assertEquals("23,232,323.0000",excelCell.getShowText() );
	}
	/**
	 * 测试设置日期
	 */
	@Test
	public void testSetTime(){
		ExcelCell excelCell = new ExcelCell();
		excelCell.setText("2016/04/28");
		CellFormateUtil.setTime(excelCell, "yyyy/MM/dd");
		System.out.println(excelCell.getValue());
		System.out.println(excelCell.getShowText());
		excelCell.setText("2016年04月28日");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd日");
		excelCell.setText("2016年0428");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd日");
		excelCell.setText("20160428日");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd日");
		excelCell.setText("20160428日");
		CellFormateUtil.setTime(excelCell, "yyyy年MMdd日");
		excelCell.setText("20160428日");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd");
		excelCell.setText("201604月28");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd日");
		excelCell.setText("20160428");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd日");
		//System.out.println(excelCell.getValue());
		excelCell.setText("2016年04月");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月");
		excelCell.setText("2016年04月");
		CellFormateUtil.setTime(excelCell, "yyyy年MM");
		excelCell.setText("2016年13月");
		CellFormateUtil.setTime(excelCell, "yyyy年MM");
		excelCell.setText("2016年13月1日");
		CellFormateUtil.setTime(excelCell, "yyyy年MM月dd");
		excelCell.setText("2016/13/1");
		CellFormateUtil.setTime(excelCell, "yyyy/MM/dd");
		excelCell.setText("2016/13/1");
		CellFormateUtil.setTime(excelCell, "");
		excelCell.setText("2016/12/1");
		CellFormateUtil.setTime(excelCell, "");
		excelCell.setText("2016年12月1日");
		CellFormateUtil.setTime(excelCell, "");
		//System.out.println(excelCell.getValue());
	}
	/**
	 * 测试货币
	 */
	@Test
	public void testSetCurrency() {
		// 小数设置4位小数点和千分位
		ExcelCell excelCell = new ExcelCell();
		excelCell.setText("1231231.23");
		CellFormateUtil.setCurrency(excelCell, 4, "$");
		Assert.assertEquals("$1,231,231.2300", excelCell.getValue());
		// 整数设置4位小数点和千分位
		excelCell.setText("23232323");
		CellFormateUtil.setCurrency(excelCell, 4, "$");
		Assert.assertEquals("$23,232,323.0000", excelCell.getValue());
	}
	
	/**
	 * 测试设置百分比
	 */
	@Test
	public void testSetPercent(){
		ExcelCell excelCell = new ExcelCell();
		excelCell.setText("1.9");
		CellFormateUtil.setPercent(excelCell, 2);
		Assert.assertEquals("100.00%", excelCell.getValue());
	}
	/**
	 * 测试设置文本
	 */
	@Test
	public void testSetText(){
		ExcelCell excelCell = new ExcelCell();
		excelCell.setText("张三");
		CellFormateUtil.setText(excelCell);
		Assert.assertEquals("张三", excelCell.getValue());
	}
	
	/**
	 * 测试设置数字格式
	 */
	@Test
	public void testGetNumDataFormate(){
		//整数有千分位
		Assert.assertEquals("#,##0_);\\(#,##0\\)", CellFormateUtil.getNumDataFormate(0, true));
		//整数无千分位
		Assert.assertEquals("0_);\\(0\\)", CellFormateUtil.getNumDataFormate(0, false));
		//1位小数有千分位
		Assert.assertEquals("#,##0.0_);\\(#,##0.0\\)", CellFormateUtil.getNumDataFormate(1, true));
		//1位小数无千分位
		Assert.assertEquals("0.0_);\\(0.0\\)", CellFormateUtil.getNumDataFormate(1, false));
	}
	
	/**
	 * 测试货币格式
	 */
	@Test
	public void testGetCurrencyDataFormate() {
		// 整数
		Assert.assertEquals("\"$\"#,##0_);\\(\"$\"#,##0\\)",CellFormateUtil.getCurrencyDataFormate(0, "$"));
		// 小数
		Assert.assertEquals("\"$\"#,##0.000_);\\(\"$\"#,##0.000\\)",CellFormateUtil.getCurrencyDataFormate(3, "$"));
	}

	/**
	 * 测试百分比格式
	 */
	@Test
	public void testGetPercentDataFormate() {
		// 整数
		Assert.assertEquals("0%", CellFormateUtil.getPercentDataFormate(0));
		// 小数
		Assert.assertEquals("0.000%", CellFormateUtil.getPercentDataFormate(3));
	}
	
	/**
	 * 测试自动识别
	 */
	@Test
	public void testAutoRecognise() {
		ExcelCell excelCell = new ExcelCell();
//		CellFormateUtil.autoRecognise("123", excelCell);
//		Assert.assertEquals( CELLTYPE.NUMERIC, excelCell.getType());
//		Assert.assertEquals( 123.0, excelCell.getValue());
//		CellFormateUtil.autoRecognise("2015年5月5日", excelCell);
//		Assert.assertEquals( CELLTYPE.DATE, excelCell.getType());
//		CellFormateUtil.autoRecognise("2015年5月", excelCell);
//		Assert.assertEquals( CELLTYPE.DATE, excelCell.getType());
//		CellFormateUtil.autoRecognise("2015/5/5", excelCell);
//		Assert.assertEquals( CELLTYPE.DATE, excelCell.getType());
//		CellFormateUtil.autoRecognise("2015\5\5", excelCell);
//		Assert.assertEquals( CELLTYPE.DATE, excelCell.getType());
		CellFormateUtil.autoRecognise("1999/900/9", excelCell);
		Assert.assertEquals( CELLTYPE.BLANK, excelCell.getType());
	}
	
	@Test
	public void testSetShowText() {
//		ExcelCell excelCell = new ExcelCell();
//		Content content = new Content();
//		CustomProp customProp = new CustomProp();
//		excelCell.setType(CELLTYPE.BLANK);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.DATE);
//		excelCell.getCellstyle().setDataformat("123");
//		excelCell.setType(CELLTYPE.DATE);
//		excelCell.getCellstyle().setDataformat("General");
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("@");
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setValue(new Date());
//		excelCell.getCellstyle().setDataformat("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.DATE);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("yyyy\"年\"m\"月\";@");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.DATE);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("m/d/yy");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.DATE);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("0_);\\(0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("#,##0_);\\(#,##0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("\"¥\"#,##0_);\\(\"¥\"#,##0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("\"$\"#,##0_);\\(\"¥\"#,##0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("0%");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("0.0_);\\(0.0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("#,##0.0_);\\(#,##0.0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("\"¥\"#,##0.0_);\\(\"¥\"#,##0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("\"$\"#,##0.0_);\\(\"$\"#,##0\\)");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("0.00%");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("0.01%");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.getCellstyle().setDataformat("1.00%");
//		excelCell.setType(CELLTYPE.STRING);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
//		excelCell.setType(CELLTYPE.NUMERIC);
//		CellFormateUtil.setShowText(excelCell, content, customProp);
		
	}
}
