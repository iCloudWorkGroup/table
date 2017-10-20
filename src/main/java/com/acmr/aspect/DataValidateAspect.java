package com.acmr.aspect;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.After;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.springframework.stereotype.Component;

import com.acmr.cache.MemoryUtil;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.copy.TempObj;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.history.ChangeArea;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

@Aspect
@Component
public class DataValidateAspect {
	@After("execution(public void com.acmr.excel.service.CellService.addRow(..))")
	public void addRow(JoinPoint joinPoint) {
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		RowOperate rowOperate = (RowOperate) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Integer> rowMap = data.getRowMap();
		
		String upRowCode = excelSheet.getRows().get(rowOperate.getRow() - 1).getCode();
		Integer rule = rowMap.get(upRowCode);
		String rowCode = excelSheet.getRows().get(rowOperate.getRow()).getCode();
		if (rule != null) {
			rowMap.put(rowCode, rule);
		}
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		Map<String, Integer> colRuleMap = cellMap.get(upRowCode);
		if(colRuleMap != null){
			Map<String,Integer> crMap =  new HashMap<String,Integer>();
			cellMap.put(rowCode, crMap);
			for(String col : colRuleMap.keySet()){
				crMap.put(col, colRuleMap.get(col));
			}
		}
	}
	@After("execution(public void com.acmr.excel.service.CellService.deleteRow(..))")
	public void deleteRow(JoinPoint joinPoint) {
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		RowOperate rowOperate = (RowOperate) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Integer> rowMap = data.getRowMap();
		String rowCode = excelSheet.getRows().get(rowOperate.getRow()).getCode();
		rowMap.remove(rowCode);
		data.getCellMap().remove(rowCode);
	}
	@After("execution(public void com.acmr.excel.service.CellService.addCol(..))")
	public void addCol(JoinPoint joinPoint) {
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		ColOperate colOperate = (ColOperate) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Integer> colMap = data.getColMap();
		String upColCode = excelSheet.getCols().get(colOperate.getCol() - 1).getCode();
		Integer rule = colMap.get(upColCode);
		String colCode = excelSheet.getCols().get(colOperate.getCol()).getCode();
		if (rule != null) {
			colMap.put(colCode, rule);
		}
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		for(String row : cellMap.keySet()){
			Map<String, Integer> colRuleMap = cellMap.get(row);
			if(colRuleMap != null && colRuleMap.get(upColCode) != null){
				cellMap.get(row).put(colCode, colRuleMap.get(upColCode));
			}
		}
	}
	@After("execution(public void com.acmr.excel.service.CellService.deleteCol(..))")
	public void deleteCol(JoinPoint joinPoint) {
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		ColOperate colOperate = (ColOperate) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Integer> colMap = data.getColMap();
		String colCode = excelSheet.getCols().get(colOperate.getCol()).getCode();
		colMap.remove(colCode);
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		for(String row : cellMap.keySet()){
			Map<String, Integer> colRuleMap = cellMap.get(row);
			if(colRuleMap != null){
				colRuleMap.remove(colCode);
			}
		}
	}
	
	@Before("execution(public void com.acmr.excel.service.CellService.mergeCell(..))")
	public void merge(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		Cell cell = (Cell) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		int firstRow = cell.getCoordinate().getStartRow();
		int firstCol = cell.getCoordinate().getStartCol();
		int lastRow = cell.getCoordinate().getEndRow();
		int lastCol = cell.getCoordinate().getEndCol();
		String rowCode = excelSheet.getRows().get(firstRow).getCode();
		String colCode = excelSheet.getCols().get(firstCol).getCode();
		int rule = getRuleIndex(firstRow, lastRow, firstCol, lastCol, excelSheet, cellMap);
		if(rule != -1){
			Map<String, Integer> cMap = cellMap.get(rowCode);
			if(cellMap.get(rowCode) == null){
				cMap = new HashMap<String, Integer>();
				cellMap.put(rowCode, cMap);
			}
			cMap.put(colCode,rule);
		}
		for (int i = firstRow; i <= lastRow; i++) {
			String rCode = excelSheet.getRows().get(i).getCode();
			Map<String, Integer> colRuleMap = cellMap.get(rCode);
			if(colRuleMap == null) continue;
			for (int j = firstCol; j <= lastCol; j++) {
				if(i == 0 && j==0) continue;
				String cCode = excelSheet.getCols().get(j).getCode();
				colRuleMap.remove(cCode);
			}
		}
		
	}
	
	private int getRuleIndex(int firstRow,int lastRow,int firstCol,int lastCol,ExcelSheet excelSheet,Map<String, Map<String, Integer>> cellMap){
		for (int i = firstRow; i <= lastRow; i++) {
			List<ExcelCell> cellList = excelSheet.getRows().get(i).getCells();
			for (int j = firstCol; j <= lastCol; j++) {
				if(cellList.get(j) == null) continue;
				String text = cellList.get(j).getText();
				if(text != null && !"".equals(text)){
					String rowCode = excelSheet.getRows().get(i).getCode();
					String colCode = excelSheet.getCols().get(j).getCode();
					Map<String, Integer> colRuleMap = cellMap.get(rowCode);
					if(colRuleMap == null) return -1;
					return colRuleMap.get(colCode) == null ? -1 :  colRuleMap.get(colCode);
				}
			}
		}
		return -1;
	}
	
	
	@After("execution(public void com.acmr.excel.service.CellService.splitCell(..))")
	public void split(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		Cell cell = (Cell) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		int firstRow = cell.getCoordinate().getStartRow();
		int firstCol = cell.getCoordinate().getStartCol();
		int lastRow = cell.getCoordinate().getEndRow();
		int lastCol = cell.getCoordinate().getEndCol();
		String fRowCode = excelSheet.getRows().get(firstRow).getCode();
		String fColCode = excelSheet.getCols().get(firstCol).getCode();
		Map<String, Integer> colRuleMap = cellMap.get(fRowCode);
		if(colRuleMap == null) return;
		Integer rule = colRuleMap.get(fColCode);
		if(rule == null) return;
		for (int i = firstRow; i <= lastRow; i++) {
			String rowCode = excelSheet.getRows().get(i).getCode();
			colRuleMap = cellMap.get(rowCode);
			if(colRuleMap == null) {
				colRuleMap = new HashMap<String, Integer>();
				cellMap.put(rowCode, colRuleMap);
			}
			for (int j = firstCol; j <= lastCol; j++) {
				String colCode = excelSheet.getCols().get(j).getCode();
				colRuleMap.put(colCode, rule);
			}
		}
	}
	@After("execution(public void com.acmr.excel.service.PasteService.copy(..))")
	public void copy(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		Copy copy = (Copy) args[0];
		ExcelBook excelBook = (ExcelBook) args[1];
		String excelId = (String) args[2];
		copyOrCut(copy, excelBook, null, excelId);
		
	}
	@After("execution(public void com.acmr.excel.service.PasteService.cut(..))")
	public void cut(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		Copy copy = (Copy) args[0];
		ExcelBook excelBook = (ExcelBook) args[1];
		String excelId = (String) args[2];
		copyOrCut(copy, excelBook, "cut", excelId);
		
	}
	private void copyOrCut(Copy copy,ExcelBook excelBook,String flag,String excelId){
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int targetRowIndex = copy.getTarget().getOprRow();
		int targetColIndex = copy.getTarget().getOprCol();
		Data data = getData(excelId);
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		List<ExcelColumn> colList = excelBook.getSheets().get(0).getCols();
		List<TempObj> temList = new ArrayList<TempObj>();
		for (int i = startRowIndex; i <= endRowIndex; i++) {
			String rowCode = rowList.get(i).getCode();
			Map<String, Integer> colRuleMap = cellMap.get(rowCode);
			if(colRuleMap == null) continue;
			int tempColIndex = targetColIndex;
			for (int j = startColIndex; j <= endColIndex; j++) {
				String colCode = colList.get(j).getCode();
				Integer rule = colRuleMap.get(colCode);
				if(rule == null) continue;
				TempObj tempObj = new TempObj();
				tempObj.setRow(targetRowIndex);
				tempObj.setCol(tempColIndex);
				tempObj.setRule(rule);
				temList.add(tempObj);
				if ("cut".equals(flag)) {
					colRuleMap.remove(colCode);
				}
				tempColIndex++;
			}
			targetRowIndex++;
		}

		for (TempObj tempObj : temList) {
			String rowCode = rowList.get(tempObj.getRow()).getCode();
			Map<String, Integer> colRuleMap = cellMap.get(rowCode);
			if(colRuleMap == null){
				colRuleMap = new HashMap<String, Integer>();
				cellMap.put(rowCode, colRuleMap);
			}
			String colCode = colList.get(tempObj.getCol()).getCode();
			colRuleMap.put(colCode, tempObj.getRule());
		}
	}
	
	private Data getData(String excelId) {
		Data data = MemoryUtil.getDataValidateMap().get(excelId);
		if (data == null) {
			data = new Data();
			MemoryUtil.getDataValidateMap().put(excelId, data);
		}
		return data;
	}
	
	
}
