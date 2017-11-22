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
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.copy.TempObj;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.Rule;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.util.StringUtil;

import acmr.excel.ExcelHelper;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

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
	@Before("execution(public void com.acmr.excel.service.CellService.deleteRow(..))")
	public void deleteRow(JoinPoint joinPoint) {
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		RowOperate rowOperate = (RowOperate) args[1];
		String excelId = (String) args[2];
		Data data = getData(excelId);
		Map<String, Integer> rowMap = data.getRowMap();
		ExcelRow excelRow = excelSheet.getRows().get(rowOperate.getRow());
		String rowCode = excelRow.getCode();
		rowMap.remove(rowCode);
		data.getCellMap().remove(rowCode);
		updateRule(excelRow.getExps(), data.getRuleList(), excelSheet, rowCode);
	}
	
	
	private void updateRule(Map<String,String> exps,List<Rule> rules,ExcelSheet excelSheet,String code){
		String rule = exps.get("rule");
		if (!StringUtil.isEmpty(rule)) {
			String[] ruleList = rule.split(",");
			for(String rl : ruleList){
				Rule r = rules.get(Integer.valueOf(rl));
				String[] position = r.getFormula1().split(":");
				String start = position[0];
				String end = position[1];
				ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
				ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
				if (start.contains("R") && start.contains("C") && end.contains("R") && end.contains("C")) {
					String aliasColStart = start.substring(1,start.lastIndexOf("R"));
					String aliasRowStart = start.substring(start.lastIndexOf("R") + 1);
					int colStart = colList.getMaps().get(aliasColStart);
					int rowStart = rowList.getMaps().get(aliasRowStart);
					String aliasColEnd = end.substring(1, end.lastIndexOf("R"));
					String aliasRowEnd = end.substring(end.lastIndexOf("R") + 1);
					int colEnd = colList.getMaps().get(aliasColEnd);
					int rowEnd = rowList.getMaps().get(aliasRowEnd);
					if(code.equals(aliasRowStart)){
						updateExpress(rowList.get(rowStart + 1).getExps(), rl, 
								"C"+aliasColStart+"R"+rowList.get(rowStart+1).getCode()+":C"+aliasColEnd+"R"+aliasRowEnd,r);
					}else if(code.equals(aliasRowEnd)){
						updateExpress(rowList.get(rowEnd-1).getExps(), rl, "C"+aliasColStart+"R"+aliasRowStart+":C"+aliasColEnd+
								"R"+rowList.get(rowEnd-1).getCode(), r);
					}else if(code.equals(aliasColStart)){
						updateExpress(colList.get(colStart+1).getExps(), rl, "C"+colList.get(colStart+1).getCode()+"R"+
									aliasRowStart+":C"+aliasColEnd+"R"+aliasRowEnd, r);
					}else if(code.equals(aliasColEnd)){
						updateExpress(colList.get(colEnd-1).getExps(), rl, "C"+aliasColStart+"R"+aliasRowStart+":C"+
						       colList.get(colEnd-1).getCode()+"R"+aliasRowEnd, r);
					}
				}
			}
		}
	}
	/**
	 * 更新表达式
	 * @param map
	 * @param rl
	 * @param express
	 * @param r
	 */
	private void updateExpress(Map<String, String> map,String rl,String express,Rule r){
		String rList = map.get("rule");
		if (StringUtil.isEmpty(rList)) {
			rList = "";
		}
		String newList = rl+ "," + rList;
		map.put("rule", newList);
		r.setFormula1(express);
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
	@Before("execution(public void com.acmr.excel.service.CellService.deleteCol(..))")
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
		updateRule(excelSheet.getCols().get(colOperate.getCol()).getExps(), data.getRuleList(), excelSheet, colCode);
	}
	
	@Before("execution(public void com.acmr.excel.service.CellService.mergeCell(..))")
	public void merge(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		Cell cell = (Cell) args[1];
		String excelId = (String) args[2];
		VersionHistory versionHistory = (VersionHistory)args[3];
		int step = (int)args[4];
		versionHistory.getVersion().put(step, step);
		History history = new History();
		history.setOperatorType(OperatorConstant.merge);
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
			ChangeArea changeArea = new ChangeArea();
			changeArea.setRowIndex(firstRow);
			changeArea.setColIndex(firstCol);
			changeArea.setType(1);
			changeArea.getOriginalValues().set(1, cMap.get(colCode));
			cMap.put(colCode,rule);
			changeArea.getUpdateValues().set(1, cMap.get(colCode));
			history.getChangeAreaList().add(changeArea);
		}
		for (int i = firstRow; i <= lastRow; i++) {
			String rCode = excelSheet.getRows().get(i).getCode();
			Map<String, Integer> colRuleMap = cellMap.get(rCode);
			if(colRuleMap == null) continue;
			for (int j = firstCol; j <= lastCol; j++) {
				if(i == 0 && j==0) continue;
				ChangeArea changeArea = new ChangeArea();
				changeArea.setRowIndex(i);
				changeArea.setColIndex(j);
				changeArea.setType(1);
				String cCode = excelSheet.getCols().get(j).getCode();
				changeArea.getOriginalValues().set(1, colRuleMap.get(cCode));
				colRuleMap.remove(cCode);
				changeArea.getUpdateValues().set(1, null);
				history.getChangeAreaList().add(changeArea);
			}
		}
		versionHistory.getMap().put(step, history);
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
	
	
	@Before("execution(public void com.acmr.excel.service.CellService.splitCell(..))")
	public void split(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		ExcelSheet excelSheet = (ExcelSheet) args[0];
		Cell cell = (Cell) args[1];
		String excelId = (String) args[2];
		VersionHistory versionHistory = (VersionHistory)args[3];
		int step = (int)args[4];
		versionHistory.getVersion().put(step, step);
		History history = new History();
		history.setOperatorType(OperatorConstant.mergedelete);
		Data data = getData(excelId);
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		int firstRow = cell.getCoordinate().getStartRow();
		int firstCol = cell.getCoordinate().getStartCol();
		int lastRow = cell.getCoordinate().getEndRow();
		int lastCol = cell.getCoordinate().getEndCol();
		String fRowCode = excelSheet.getRows().get(firstRow).getCode();
		String fColCode = excelSheet.getCols().get(firstCol).getCode();
		Map<String, Integer> colRuleMap = cellMap.get(fRowCode);
		versionHistory.getMap().put(step, history);
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
				ChangeArea changeArea = new ChangeArea();
				changeArea.setRowIndex(i);
				changeArea.setColIndex(j);
				changeArea.setType(1);
				changeArea.getOriginalValues().set(1, colRuleMap.get(colCode));
				colRuleMap.put(colCode, rule);
				changeArea.getUpdateValues().set(1, rule);
				history.getChangeAreaList().add(changeArea);
			}
		}
		versionHistory.getMap().put(step, history);
	}
	@After("execution(public void com.acmr.excel.service.PasteService.copy(..))")
	public void copy(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		Copy copy = (Copy) args[0];
		ExcelBook excelBook = (ExcelBook) args[1];
		String excelId = (String) args[2];
		VersionHistory versionHistory = (VersionHistory)args[3];
		int step = (int)args[4];
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
		}
		version += 1;
		History history = versionHistory.getMap().get(version);
		history.setOperatorType(OperatorConstant.copy);
		copyOrCut(copy, excelBook, null, excelId,history);
		
	}
	@After("execution(public void com.acmr.excel.service.PasteService.cut(..))")
	public void cut(JoinPoint joinPoint){
		Object[] args = joinPoint.getArgs();
		Copy copy = (Copy) args[0];
		ExcelBook excelBook = (ExcelBook) args[1];
		String excelId = (String) args[2];
		VersionHistory versionHistory = (VersionHistory)args[3];
		int step = (int)args[4];
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
		}
		version += 1;
		History history = versionHistory.getMap().get(version);
		history.setOperatorType(OperatorConstant.cut);
		copyOrCut(copy, excelBook, "cut", excelId,history);
		
	}
	private void copyOrCut(Copy copy,ExcelBook excelBook,String flag,String excelId,History history){
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
					ChangeArea changeArea = new ChangeArea();
					changeArea.setColIndex(i);
					changeArea.setRowIndex(j);
					changeArea.setType(1);
					changeArea.getOriginalValues().set(1, rule);
					colRuleMap.remove(colCode);
					changeArea.getUpdateValues().set(1, null);
					history.getChangeAreaList().add(changeArea);
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
			ChangeArea changeArea = new ChangeArea();
			changeArea.setColIndex(tempObj.getCol());
			changeArea.setRowIndex(tempObj.getRow());
			changeArea.setType(1);
			changeArea.getOriginalValues().set(1, colRuleMap.get(colCode));
			colRuleMap.put(colCode, tempObj.getRule());
			changeArea.getUpdateValues().set(1, tempObj.getRule());
			history.getChangeAreaList().add(changeArea);
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
