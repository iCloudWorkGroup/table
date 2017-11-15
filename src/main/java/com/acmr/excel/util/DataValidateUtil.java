package com.acmr.excel.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import acmr.excel.ExcelHelper;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellRangeAddress;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelDataValidation;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.datavalidate.CellList;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.Model;
import com.acmr.excel.model.datavalidate.RowColList;
import com.acmr.excel.model.datavalidate.Rule;
import com.acmr.excel.test.TestUtil;

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
		for(ExcelDataValidation excelDataValidation : excelDataValidations){
			if(excelDataValidation.getValidationType() == 3){
				excelDataValidation.setFormula1(alias2Display(excelDataValidation.getFormula1(), excelSheet));
			}
		}
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
			rule.setValidationType(excelDataValidation.getValidationType());
			rule.setFormula1(excelDataValidation.getFormula1());
			if(excelDataValidation.getValidationType() == 3){
				rule.setFormula1(display2Alias(excelDataValidation.getFormula1(), excelSheet));
			}
			rule.setFormula2(excelDataValidation.getFormula2());
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
	
	
	public static String display2Alias(String text,ExcelSheet excelSheet){
		if(text == null || "".equals(text)) return null;
		if(text.indexOf("=") == -1){
			return text;
		}
		String[] position = text.split(":");
		String start = position[0];
		String end = position[1];
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		String newText = null;
		if (start.indexOf("$") != start.lastIndexOf("$") && end.indexOf("$") != end.lastIndexOf("$")) {
			String displayColStart = start.substring(1, start.lastIndexOf("$"));
			String displayRowStart = start.substring(start.lastIndexOf("$")+1);
			String aliasColStart = colList.get(ExcelHelper.getColIndex(displayColStart)).getCode();
			String aliasRowStart = rowList.get(Integer.valueOf(displayRowStart)-1).getCode();

			String displayColEnd = end.substring(1, end.lastIndexOf("$"));
			String displayRowEnd = end.substring(end.lastIndexOf("$")+1);
			String aliasColEnd = colList.get(ExcelHelper.getColIndex(displayColEnd)).getCode();
			String aliasRowEnd = rowList.get(Integer.valueOf(displayRowEnd)-1).getCode();
			newText = "C" + aliasColStart + "R" + aliasRowStart + ":C" + aliasColEnd + "R" + aliasRowEnd;
		}else {
			String displayStart = start.substring(2);
			//String displayEnd = end.substring(1);
			if(isNumeric(displayStart)){ 
				//newText = "$"+rowList.get(Integer.valueOf(displayStart)-1).getCode() + ":$" + rowList.get(Integer.valueOf(displayEnd)-1).getCode();
				int value = rowList.getMaps().get(displayStart) + 1;
				newText = "R" + value + ":R" + value;
			}else{
				//newText = "$" + colList.get(ExcelHelper.getColIndex(displayStart)).getCode() + ":$" + colList.get(ExcelHelper.getColIndex(displayEnd)).getCode();
				String value = colList.get(ExcelHelper.getColIndex(displayStart)).getCode();
				newText = "C" + value + ":C" + value;
			}
		}
		return newText;
	} 
	public static String alias2Display(String text,ExcelSheet excelSheet){
		if (text == null || "".equals(text))
			return null;
		if (!text.contains("R") && !text.contains("C")) return text;
		String[] position = text.split(":");
		String start = position[0];
		String end = position[1];
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		String newText = null;
		if (start.contains("R") && start.contains("C") && end.contains("R") && end.contains("C")) {
			String aliasColStart = start.substring(1, start.lastIndexOf("R"));
			String aliasRowStart = start.substring(start.lastIndexOf("R")+1);
			String displayColStart = ExcelHelper.getColCode(colList.getMaps().get(aliasColStart));
			String displayRowStart = rowList.getMaps().get(aliasRowStart) + 1 +"";

			String aliasColEnd = end.substring(1, end.lastIndexOf("R"));
			String aliasRowEnd = end.substring(end.lastIndexOf("R")+1);
			String displayColEnd = ExcelHelper.getColCode(colList.getMaps().get(aliasColEnd));
			String displayRowEnd = rowList.getMaps().get(aliasRowEnd) + 1 +"";
			newText = "$" + displayColStart + "$" + displayRowStart + ":$" + displayColEnd + "$" + displayRowEnd;
		}else {
			String aliasCode = start.substring(1);
			String display = null;
			if(start.contains("R")){ 
				display = rowList.getMaps().get(aliasCode) + 1 + "";
			}else{
				display = ExcelHelper.getColCode(colList.getMaps().get(aliasCode));
			}
			newText = "$"+display + ":$" + display; 
		}
		newText = "="+newText;
		return newText;
	} 

	public static boolean isNumeric(String str) {
		Pattern pattern = Pattern.compile("[0-9]*");
		Matcher isNum = pattern.matcher(str);
		if (!isNum.matches()) {
			return false;
		}
		return true;
	}

	
	
	
	
	
	
	
	
	
	
	
	
	
	public static ExcelBook createNewExcel() {
		ExcelBook excelBook = new ExcelBook();
		ExcelSheet sheet = new ExcelSheet();
		for (int i = 1; i < 27; i++) {
			ExcelColumn column = sheet.addColumn();
			column.setWidth(69);
		}
		for (int i = 1; i < 101; i++) {
			ExcelRow row = sheet.addRow();
			row.setHeight(19);
		}
		excelBook.getSheets().add(sheet);
		return excelBook;
	}
	public static void main(String[] args) {
		
		System.out.println(display2Alias("=$A$1:$A$6", createNewExcel().getSheets().get(0)));
		System.out.println(display2Alias("=$C:$C", createNewExcel().getSheets().get(0)));
		System.out.println(display2Alias("=$3:$3", createNewExcel().getSheets().get(0)));
		System.out.println(alias2Display("C1R1:C1R6", createNewExcel().getSheets().get(0)));
		System.out.println(alias2Display("R3:R3", createNewExcel().getSheets().get(0)));
		System.out.println(alias2Display("C3:C3", createNewExcel().getSheets().get(0)));
	}
	
	
	
}
