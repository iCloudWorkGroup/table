package com.acmr.excel.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellRangeAddress;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelDataValidation;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.datavalidate.CellList;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.RowColList;
import com.acmr.excel.model.datavalidate.Rule;

public class DataValidateUtil {
	private static final int MIN = 0; 
	private static final int ROWMAX = 1048575; 
	private static final int COLMAX = 16383; 
	
	/**
	 * map转list
	 * @param data
	 * @param excelDataValidations
	 * @param excelSheet
	 */
	public static void map2List(Data data, List<ExcelDataValidation> excelDataValidations,ExcelSheet excelSheet) {
		if(data == null || excelDataValidations == null || excelSheet == null ) return;
		Map<Integer,List<RowColList>> row = mergeRC(data.getRowMap(), (ListHashMap<ExcelRow>)excelSheet.getRows());
		Map<Integer,List<RowColList>> col = mergeRC(data.getColMap(), (ListHashMap<ExcelColumn>)excelSheet.getCols());
		Map<Integer, List<CellList>> cell = mergeCellRange(data.getCellMap(), (ListHashMap<ExcelRow>)excelSheet.getRows(), (ListHashMap<ExcelColumn>)excelSheet.getCols());
		setRowValidation(row, excelDataValidations,data.getRuleList());
		setColValidation(col, excelDataValidations,data.getRuleList());
		setCellValidation(cell, excelDataValidations,data.getRuleList());
	}
	/**
	 * list转map
	 * @param data
	 * @param excelDataValidations
	 * @param excelSheet
	 */
	public static void list2Map(Data data, List<ExcelDataValidation> excelDataValidations,ExcelSheet excelSheet) {
		Map<String, Integer> rowMap = data.getRowMap();
		Map<String, Integer> colMap = data.getColMap();
		Map<String, Map<String, Integer>> cellMap = data.getCellMap();
		List<Rule> ruleList = data.getRuleList();
		for(ExcelDataValidation excelDataValidation : excelDataValidations){
			Rule rule = new Rule();
			rule.setFormula1(excelDataValidation.getFormula1());
			rule.setFormula2(excelDataValidation.getFormula2());
			rule.setValidationType(excelDataValidation.getValidationType());
			int ruleIndex = ruleList.indexOf(rule);
			if (ruleIndex == -1) {
				ruleList.add(rule);
				ruleIndex = ruleList.size() - 1;
			}
			List<ExcelCellRangeAddress> excelCellRangeAddresses = excelDataValidation.getExcelCellRangeAddresses();
			for (ExcelCellRangeAddress excelCellRangeAddress : excelCellRangeAddresses) {
				int firstCol = excelCellRangeAddress.getFirstColumn();
				int firstRow = excelCellRangeAddress.getFirstRow();
				int lastCol = excelCellRangeAddress.getLastColumn();
				int lastRow = excelCellRangeAddress.getLastRow();
				if (firstCol == MIN && lastCol == COLMAX) {
					for (int i = firstRow; i <= lastRow; i++) {
						ExcelRow excelRow = excelSheet.getRows().get(i);
						rowMap.put(excelRow.getCode(), ruleIndex);
					}
				} else if (firstRow == MIN && lastRow == ROWMAX) {
					for (int i = firstCol; i <= lastCol; i++) {
						ExcelColumn excelColumn = excelSheet.getCols().get(i);
						colMap.put(excelColumn.getCode(), ruleIndex);
					}
				} else {
					for (int i = firstRow; i <= lastRow; i++) {
						ExcelRow excelRow = excelSheet.getRows().get(i);
						Map<String, Integer> colRuleMap = cellMap.get(excelRow.getCode());
						if (colRuleMap == null) {
							colRuleMap = new HashMap<String, Integer>();
							cellMap.put(excelRow.getCode(), colRuleMap);
						}
						for (int j = firstCol; j <= lastCol; j++) {
							ExcelColumn excelColumn = excelSheet.getCols().get(j);
							colRuleMap.put(excelColumn.getCode(), ruleIndex);
						}
					}
				}
			}
		}
	}
	
	
	/**
	 * 行列合并
	 * @param map
	 * @param list
	 * @return
	 */
	
	private static Map<Integer,List<RowColList>> mergeRC(Map<String, Integer> map,ListHashMap list){
		Map<Integer,List<RowColList>> retMap = new HashMap<Integer,List<RowColList>>();
		Map<Integer,TreeSet<Integer>> tempMap = new HashMap<Integer,TreeSet<Integer>>();
		for (String key : map.keySet()) {
			Integer rule = map.get(key);
			TreeSet<Integer> tset = tempMap.get(rule);
			if(tset == null){
				tset = new TreeSet<Integer>();
				tempMap.put(rule, tset);
			}
			tset.add((Integer)list.getMaps().get(key));
		}
		for(Integer rule : tempMap.keySet()){
			List<RowColList> retList = new ArrayList<RowColList>();
			retMap.put(rule, retList);
			TreeSet<Integer> set = tempMap.get(rule);
			List<Integer> tempList = new ArrayList<Integer>(set);
//			for(int i = 0;i<tempList.size();){
//				int current = tempList.get(i);
//				RowColList rowColList = new RowColList();
//				rowColList.setStart(current);
//				for(int j = i+1 ;j<tempList.size();j++){
//					if( current +(j-i) != tempList.get(j)){
//						rowColList.setEnd(tempList.get(i));
//						i++;
//						break;
//					}else{
//						rowColList.setEnd(tempList.get(j));
//					}
//					i = j+1;
//				}
//				if(tempList.size() ==1 ){
//					i++;
//				}
//				retList.add(rowColList);
//			}
			for(int i = 0;i<tempList.size();i++){
				RowColList rowColList = new RowColList();
				rowColList.setStart(tempList.get(i));
				rowColList.setEnd(tempList.get(i));
				retList.add(rowColList);
			}
		}
		return retMap;
	}
	
	
	private static Map<Integer,List<CellList>> mergeCellRange(Map<String, Map<String, Integer>> cellMap,ListHashMap<ExcelRow> rowlist,ListHashMap<ExcelColumn> collist){
		Map<Integer,List<CellList>> retMap = new HashMap<Integer,List<CellList>>();
		Map<Integer,TreeMap<Integer,TreeSet<Integer>>> tempMap = new HashMap<Integer,TreeMap<Integer,TreeSet<Integer>>>();
		for(String row : cellMap.keySet()){
			Map<String, Integer> colRuleMap = cellMap.get(row);
			for(String col : colRuleMap.keySet()){
				Integer rule = colRuleMap.get(col);
				TreeMap<Integer, TreeSet<Integer>> treMap = tempMap.get(rule);
				if(treMap == null){
					treMap = new TreeMap<Integer, TreeSet<Integer>>();
					tempMap.put(rule, treMap);
				}
				Integer rowIndex = rowlist.getMaps().get(row);
				TreeSet<Integer> treSet = treMap.get(rowIndex);
				if(treSet == null){
					treSet = new TreeSet<Integer>();
					treMap.put(rowIndex, treSet);
				}
				treSet.add((Integer)collist.getMaps().get(col));
			}
		}
		
		
		for (Integer rule : tempMap.keySet()) {
			List<CellList> retList = new ArrayList<CellList>();
			retMap.put(rule, retList);
			TreeMap<Integer, TreeSet<Integer>> rowMap = tempMap.get(rule);
			for (Integer row : rowMap.keySet()) {
				TreeSet<Integer> currentColSet = rowMap.get(row);
				if(currentColSet == null) continue;
				//TreeSet<Integer> nextColSet = rowMap.get(k+1);
				List<Integer> currentColList = new ArrayList<>(currentColSet);
				List<Integer> nextColList = new ArrayList<>(currentColSet);
				for (int i = 0; i < currentColList.size();i++) {
					int current = currentColList.get(i);
					ExcelCell excelCell = rowlist.get(row).getCells().get(current);
					int colspan = 1;
					int rowspan = 1;
					if(excelCell != null){
						colspan = excelCell.getColspan();
						rowspan = excelCell.getRowspan();
					}
					CellList cellList = new CellList();
					cellList.setFirstColumn(current);
					cellList.setFirstRow(row);
					cellList.setLastRow(row);
					cellList.setLastColumn(current);
//					for(int j = i+1 ;j<currentColList.size();j++){
//						if(j != currentColList.size() -1 && current +(j-i)*colspan != currentColList.get(j)){
//							break;
//						}else{
//							cellList.setLastColumn(currentColList.get(j));
//						}
//						i = j+1;
//					}
					retList.add(cellList);
				}
			}
		}
		return retMap;
	}
	
	/**
	 * 设置行验证
	 * @param map
	 * @param excelDataValidations
	 */
	private static void setRowValidation(Map<Integer,List<RowColList>> map,List<ExcelDataValidation> excelDataValidations,
			List<Rule> ruleList){
		for(Integer rule : map.keySet()){
			ExcelDataValidation excelDataValidation = new ExcelDataValidation();
			Rule r = ruleList.get(rule);
			excelDataValidation.setValidationType(r.getValidationType());
			excelDataValidation.setFormula1(r.getFormula1());
			excelDataValidation.setFormula2(r.getFormula2());
			for(RowColList rowColList : map.get(rule)){
				ExcelCellRangeAddress  excelCellRangeAddress = new ExcelCellRangeAddress();
				excelCellRangeAddress.setFirstColumn(MIN);
				excelCellRangeAddress.setLastColumn(COLMAX);
				excelCellRangeAddress.setFirstRow(rowColList.getStart());
				excelCellRangeAddress.setLastRow(rowColList.getEnd());
				excelDataValidation.getExcelCellRangeAddresses().add(excelCellRangeAddress);
			}
			excelDataValidations.add(excelDataValidation);
		}
	}
	/**
	 * 设置列验证
	 * @param map
	 * @param excelDataValidations
	 */
	private static void setCellValidation(Map<Integer,List<CellList>> map,List<ExcelDataValidation> excelDataValidations,
			List<Rule> ruleList){
		for(Integer rule : map.keySet()){
			ExcelDataValidation excelDataValidation = new ExcelDataValidation();
			Rule r = ruleList.get(rule);
			excelDataValidation.setValidationType(r.getValidationType());
			excelDataValidation.setFormula1(r.getFormula1());
			excelDataValidation.setFormula2(r.getFormula2());
			for(CellList cellList : map.get(rule)){
				ExcelCellRangeAddress  excelCellRangeAddress = new ExcelCellRangeAddress();
				excelCellRangeAddress.setFirstColumn(cellList.getFirstColumn());
				excelCellRangeAddress.setLastColumn(cellList.getLastColumn());
				excelCellRangeAddress.setFirstRow(cellList.getFirstRow());
				excelCellRangeAddress.setLastRow(cellList.getLastRow());
				excelDataValidation.getExcelCellRangeAddresses().add(excelCellRangeAddress);
			}
			excelDataValidations.add(excelDataValidation);
		}
	}
	/**
	 * 设置单元格验证
	 * @param map
	 * @param excelDataValidations
	 */
	private static void setColValidation(Map<Integer,List<RowColList>> map,List<ExcelDataValidation> excelDataValidations,
			List<Rule> ruleList){
		for(Integer rule : map.keySet()){
			ExcelDataValidation excelDataValidation = new ExcelDataValidation();
			Rule r = ruleList.get(rule);
			excelDataValidation.setValidationType(r.getValidationType());
			excelDataValidation.setFormula1(r.getFormula1());
			excelDataValidation.setFormula2(r.getFormula2());
			for(RowColList rowColList : map.get(rule)){
				ExcelCellRangeAddress  excelCellRangeAddress = new ExcelCellRangeAddress();
				excelCellRangeAddress.setFirstColumn(rowColList.getStart());
				excelCellRangeAddress.setLastColumn(rowColList.getEnd());
				excelCellRangeAddress.setFirstRow(MIN);
				excelCellRangeAddress.setLastRow(ROWMAX);
				excelDataValidation.getExcelCellRangeAddresses().add(excelCellRangeAddress);
			}
			excelDataValidations.add(excelDataValidation);
		}
	}
}
