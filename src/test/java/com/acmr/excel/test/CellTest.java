package com.acmr.excel.test;

import java.util.ArrayList;
import java.util.List;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;

import com.acmr.excel.model.AreaSet;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;

/**
 * 单元格操作测试类
 * 
 * @author jinhr
 *
 */
public class CellTest {
	private HandleExcelService handleExcelService;
	private ExcelBook excelBook;

	@Before
	public void before() {
		handleExcelService = new HandleExcelService();
		excelBook = TestUtil.createNewExcel();
	}

	/**
	 * 测试批量设置单元格颜色
	 */
	@Test
	public void testBatchColorSet() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		Coordinate coordinate2 = new Coordinate();
		coordinate2.setEndRow(5);
		coordinate2.setStartRow(3);
		coordinate2.setStartCol(0);
		coordinate2.setEndCol(2);
		coordinates.add(coordinate1);
		coordinates.add(coordinate2);
		areaSet.setCoordinate(coordinates);
		String color = "rgb(73, 68, 41)";
		areaSet.setColor(color);
		handleExcelService.areaSet(areaSet, excelBook,OperatorConstant.batchcolorset);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		for (Coordinate coordinate : coordinates) {
			for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
				for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
					ExcelCell excelCell = rowList.get(i).getCells().get(j);
					ExcelColor excelColor = excelCell.getCellstyle().getFgcolor();
					Assert.assertEquals(73, excelColor.getR());
					Assert.assertEquals(68, excelColor.getG());
					Assert.assertEquals(41, excelColor.getB());
				}
			}
		}
	}
	
	/**
	 * 测试区域清除
	 */
	@Test
	public void testCleanData() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		List<ExcelRow> oriRowList = excelBook.getSheets().get(0).getRows();
		ExcelCell excelCell1 = new ExcelCell();
		excelCell1.setText("1");
		excelCell1.setValue("1");
		excelCell1.setType(CELLTYPE.STRING);
		oriRowList.get(0).set(0, excelCell1);
		ExcelCell excelCell2 = new ExcelCell();
		excelCell2.setText("2");
		excelCell2.setValue("2");
		excelCell2.setType(CELLTYPE.STRING);
		oriRowList.get(2).set(2, excelCell1);
		handleExcelService.areaSet(areaSet, excelBook,OperatorConstant.CLEANDATA);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		for (Coordinate coordinate : coordinates) {
			for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
				for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
					ExcelCell excelCell = rowList.get(i).getCells().get(j);
					if((i == coordinate.getStartRow() && j == coordinate.getStartCol()) || 
							(i == coordinate.getEndRow() && j == coordinate.getEndCol())){
						Assert.assertEquals("", excelCell.getText());
						Assert.assertNull(excelCell.getValue());
					}else{
						Assert.assertNull(excelCell);
					}
				}
			}
		}
	}
	
	/**
	 * 测试下划线
	 */
	@Test
	public void testUnderline(){
		Cell cell = new Cell();
		Coordinate coordinate = cell.getCoordinate();
		coordinate.setStartRow(0);
		coordinate.setEndRow(2);
		coordinate.setStartCol(0);
		coordinate.setEndCol(2);
		cell.setUnderline(1);
		handleExcelService.updateCells(CellUpdateType.font_underline, cell,excelBook,new VersionHistory(),1);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
			for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
				ExcelCell excelCell = rowList.get(i).getCells().get(j);
				ExcelFont excelFont = excelCell.getCellstyle().getFont();
				Assert.assertEquals(1, excelFont.getUnderline());
			}
		}
		coordinate = cell.getCoordinate();
		coordinate.setStartRow(0);
		coordinate.setEndRow(2);
		coordinate.setStartCol(0);
		coordinate.setEndCol(2);
		cell.setUnderline(0);
		handleExcelService.updateCells(CellUpdateType.font_underline, cell,excelBook,new VersionHistory(),1);
		rowList = excelBook.getSheets().get(0).getRows();
		for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
			for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
				ExcelCell excelCell = rowList.get(i).getCells().get(j);
				ExcelFont excelFont = excelCell.getCellstyle().getFont();
				Assert.assertEquals(0, excelFont.getUnderline());
			}
		}
	}
	
	@After
	public void after() {
		excelBook = null;
		handleExcelService = null;
	}
}
