package com.acmr.excel.test;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;

import com.acmr.cache.MemoryUtil;
import com.acmr.excel.model.AreaSet;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Protect;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.Rule;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.service.SheetService;

/**
 * 单元格操作测试类
 * 
 * @author jinhr
 *
 */
public class CellTest {
	private HandleExcelService handleExcelService;
	private ExcelBook excelBook;
	private SheetService sheetService;

	@Before
	public void before() {
		handleExcelService = new HandleExcelService();
		excelBook = TestUtil.createNewExcel();
		sheetService = new SheetService();
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
	/**
	 * 测试一整列变色
	 */
	@Test
	public void testColumnColor(){
		Cell cell = new Cell();
		Coordinate coordinate = cell.getCoordinate();
		coordinate.setStartRow(0);
		coordinate.setEndRow(-1);
		coordinate.setStartCol(1);
		coordinate.setEndCol(1);
		cell.setColor("rgb(147, 137, 83)");
		handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,new VersionHistory(),1);
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		ExcelColumn excelColumn = colList.get(1);
		Assert.assertEquals("rgb(147, 137, 83)", excelColumn.getExps().get("fill_bgcolor"));;
	}	
	
	/**
	 * 测试多整列变色
	 */
	@Test
	public void testMulColumnColor(){
		Cell cell = new Cell();
		Coordinate coordinate = cell.getCoordinate();
		coordinate.setStartRow(0);
		coordinate.setEndRow(-1);
		coordinate.setStartCol(1);
		coordinate.setEndCol(3);
		cell.setColor("rgb(147, 137, 83)");
		handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,new VersionHistory(),1);
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		ExcelColumn excelColumn = colList.get(1);
		ExcelColumn excelColumn2 = colList.get(2);
		ExcelColumn excelColumn3 = colList.get(3);
		Assert.assertEquals("rgb(147, 137, 83)", excelColumn.getExps().get("fill_bgcolor"));
		Assert.assertEquals("rgb(147, 137, 83)", excelColumn2.getExps().get("fill_bgcolor"));
		Assert.assertEquals("rgb(147, 137, 83)", excelColumn3.getExps().get("fill_bgcolor"));
	}
	/**
	 * 测试一整行变色
	 */
	@Test
	public void testRowColor(){
		Cell cell = new Cell();
		Coordinate coordinate = cell.getCoordinate();
		coordinate.setStartRow(1);
		coordinate.setEndRow(1);
		coordinate.setStartCol(0);
		coordinate.setEndCol(-1);
		cell.setColor("rgb(147, 137, 83)");
		handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,new VersionHistory(),1);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		ExcelRow excelRow = rowList.get(1);
		Assert.assertEquals("rgb(147, 137, 83)", excelRow.getExps().get("fill_bgcolor"));;
	}	
	
	/**
	 * 测试多整行变色
	 */
	@Test
	public void testMulRowColor(){
		Cell cell = new Cell();
		Coordinate coordinate = cell.getCoordinate();
		coordinate.setStartRow(1);
		coordinate.setEndRow(3);
		coordinate.setStartCol(0);
		coordinate.setEndCol(-1);
		cell.setColor("rgb(147, 137, 83)");
		handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,new VersionHistory(),1);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		ExcelRow excelRow = rowList.get(1);
		ExcelRow excelRow2 = rowList.get(2);
		ExcelRow excelRow3 = rowList.get(3);
		Assert.assertEquals("rgb(147, 137, 83)", excelRow.getExps().get("fill_bgcolor"));
		Assert.assertEquals("rgb(147, 137, 83)", excelRow2.getExps().get("fill_bgcolor"));
		Assert.assertEquals("rgb(147, 137, 83)", excelRow3.getExps().get("fill_bgcolor"));
	}
	
	/**
	 * 测试单元格数据验证
	 */
	@Test
	public void testCellValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "1");
		Data data = MemoryUtil.getDataValidateMap().get("1");
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		for (int i = 0; i <= 2; i++) {
			for (int j = 0; j <= 2; j++) {
				Integer r = cellMap.get(rowList.get(0).getCode()).get(colList.get(j).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
			}
		}
	}
	/**
	 * 测试行数据验证
	 */
	@Test
	public void testRowValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(-1);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "2");
		Data data = MemoryUtil.getDataValidateMap().get("2");
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		Map<String, Integer> rowMap = data.getRowMap();
		for (int i = 0; i <= 2; i++) {
				Integer r = rowMap.get(rowList.get(i).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
		}
	}
	/**
	 * 测试行数据验证
	 */
	@Test
	public void testColValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(-1);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "3");
		Data data = MemoryUtil.getDataValidateMap().get("3");
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Map<String, Integer> colMap = data.getColMap();
		for (int i = 0; i <= 2; i++) {
				Integer r = colMap.get(colList.get(i).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
		}
	}
	/**
	 * 测试行列相交数据验证
	 */
	@Test
	public void testRowColValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(-1);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "4");
		coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(-1);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "4");
		Data data = MemoryUtil.getDataValidateMap().get("4");
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		for (int i = 0; i <= 2; i++) {
			for (int j = 0; j <= 2; j++) {
				Integer r = cellMap.get(rowList.get(0).getCode()).get(colList.get(j).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
			}
		}
	}
	/**
	 * 测试列行相交数据验证
	 */
	@Test
	public void testColRowValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(-1);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "4");
		coordinate1 = new Coordinate();
		coordinate1.setEndRow(-1);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "4");
		Data data = MemoryUtil.getDataValidateMap().get("4");
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		for (int i = 0; i <= 2; i++) {
			for (int j = 0; j <= 2; j++) {
				Integer r = cellMap.get(rowList.get(0).getCode()).get(colList.get(j).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
			}
		}
	}
	/**
	 * 测试超出行数据验证
	 */
	@Test
	public void testMoreRowValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(103);
		coordinate1.setStartRow(101);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(-1);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "2");
		Data data = MemoryUtil.getDataValidateMap().get("2");
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		Map<String, Integer> rowMap = data.getRowMap();
		for (int i = 101; i <= 103; i++) {
				Integer r = rowMap.get(rowList.get(i).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
		}
	}
	/**
	 * 测试超出列数据验证
	 */
	@Test
	public void testMoreColValidate() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(-1);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(101);
		coordinate1.setEndCol(103);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		Rule rule = new Rule();
		rule.setFormula2("6");
		rule.setFormula1("4");
		rule.setValidationType(1);
		areaSet.setRule(rule);
		sheetService.dataValidate(areaSet, excelBook, "3");
		Data data = MemoryUtil.getDataValidateMap().get("3");
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Map<String, Integer> colMap = data.getColMap();
		for (int i = 101; i <= 103; i++) {
				Integer r = colMap.get(colList.get(i).getCode());
				Rule rl = data.getRuleList().get(r);
				Assert.assertEquals("6", rl.getFormula2());
				Assert.assertEquals("4", rl.getFormula1());
				Assert.assertEquals(1, rl.getValidationType());
		}
	}
	/**
	 * 测试保护
	 */
	@Test
	public void testProtect(){
		Protect protect = new Protect();
		protect.setPassword("123");
		protect.setProtect(true);
		sheetService.protect(protect, excelBook);
		Assert.assertEquals(true, excelBook.getSheets().get(0).isProtect());
		Assert.assertEquals("123", excelBook.getSheets().get(0).getPassword());
	}
	
	
	/**
	 * 测试区域单元格锁定
	 */
	@Test
	public void testAreaCellLock() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		areaSet.setLock(false);
		handleExcelService.areaSet(areaSet, excelBook,OperatorConstant.CELLLOCK);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		for (Coordinate coordinate : coordinates) {
			for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
				for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
					ExcelCell excelCell = rowList.get(i).getCells().get(j);
					Assert.assertEquals(false, excelCell.getCellstyle().isLocked());
				}
			}
		}
	}
	/**
	 * 测试区域整行锁定
	 */
	@Test
	public void testRowLock() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(3);
		coordinate1.setStartRow(2);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(-1);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		areaSet.setLock(false);
		handleExcelService.areaSet(areaSet, excelBook,OperatorConstant.CELLLOCK);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		Assert.assertEquals(false, rowList.get(2).getCellstyle().isLocked());
		Assert.assertEquals(false, rowList.get(3).getCellstyle().isLocked());
	}
	/**
	 * 测试区域整列锁定
	 */
	@Test
	public void testColLock() {
		AreaSet areaSet = new AreaSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(-1);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(2);
		coordinate1.setEndCol(3);
		coordinates.add(coordinate1);
		areaSet.setCoordinate(coordinates);
		areaSet.setLock(false);
		handleExcelService.areaSet(areaSet, excelBook,OperatorConstant.CELLLOCK);
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		Assert.assertEquals(false, colList.get(2).getCellstyle().isLocked());
		Assert.assertEquals(false, colList.get(3).getCellstyle().isLocked());
	}
	
	
	
	
	
	
	
	@After
	public void after() {
		excelBook = null;
		handleExcelService = null;
	}
}
