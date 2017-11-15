package com.acmr.excel.service;



import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.springframework.stereotype.Service;

import acmr.excel.ExcelHelper;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.util.ListHashMap;

import com.acmr.cache.MemoryUtil;
import com.acmr.excel.model.AreaSet;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Protect;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.Model;
import com.acmr.excel.model.datavalidate.Rule;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.util.DataValidateUtil;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.StringUtil;
@Service
public class SheetService {
	/**
	 * 增加行，用于初始化时向下滚动
	 * 
	 * @param sheet
	 *            SheetElement
	 * @param rowNum
	 *            增加行数
	 */
	public void addRowLine(ExcelSheet sheet, int rowNum) {
		for (int i = 0; i < rowNum; i++) {
			sheet.addRow();
		}
	}
	/**
	 * 增加列，用于初始化时向右滚动
	 * 
	 * @param sheet
	 *            SheetElement
	 * @param rowNum
	 *            增加行数
	 */
	public void addColLine(ExcelSheet sheet, int rowNum) {
		for (int i = 0; i < rowNum; i++) {
			sheet.addColumn();
		}
	}
	/**
	 * 冻结
	 * 
	 * @param excelSheet
	 *            excelSheet对象
	 * @param frozenY
	 *            冻结横坐标
	 * @param frozenX
	 *            冻结纵坐标
	 * @param startY
	 *            开始点横坐标
	 * @param startX
	 *            开始点纵坐标
	 */
	public void frozen(ExcelSheet excelSheet,Frozen frozen) {
		ExcelSheetFreeze excelSheetFreeze = excelSheet.getFreeze();
		if (excelSheetFreeze == null) {
			excelSheetFreeze = new ExcelSheetFreeze();
			excelSheet.setFreeze(excelSheetFreeze);
		}
		excelSheet.getExps().remove("fr");
		excelSheet.getExps().remove("fc");
		int frozenYIndex = frozen.getOprRow();
		int frozenXIndex = frozen.getOprCol();
		if(frozenYIndex == -1){
			frozenYIndex = 0;
			excelSheet.getExps().put("fr", "fr");
		}
		if(frozenXIndex == -1){
			frozenXIndex = 0;
			excelSheet.getExps().put("fc", "fc");
		}
		excelSheetFreeze.setRow(frozenYIndex);
		excelSheetFreeze.setCol(frozenXIndex);
		excelSheetFreeze.setFirstrow(frozenYIndex);
		excelSheetFreeze.setFirstcol(frozenXIndex);
	}
	
	public void cancelColHide(ExcelSheet excelSheet,ColOperate colOperate) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for(ExcelColumn excelColumn : colList){
			if(excelColumn.isColumnhidden()){
				excelColumn.setColumnhidden(false);
			}
		}
	}
	public void cancelRowHide(ExcelSheet excelSheet,RowOperate rowOperate) {
		List<ExcelRow> rowList = excelSheet.getRows();
		for(ExcelRow excelRow : rowList){
			if(excelRow.isRowhidden()){
				excelRow.setRowhidden(false);
			}
		}
	}
	
	
	
	public void undo(VersionHistory versionHistory,int step,ExcelSheet sheet){
		int version = versionHistory.getVersion().get(step-1);
		History hisory= versionHistory.getMap().get(version);
		int operatorType = hisory.getOperatorType();
		List<ChangeArea> changeAreaList = hisory.getChangeAreaList();
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		versionHistory.getVersion().put(step, version-1);
		switch (operatorType) {
		case OperatorConstant.textData:
		case OperatorConstant.merge:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontsize:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontfamily:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontweight:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontitalic:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.wordWrap:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fillbgcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.textDataformat:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;

		case OperatorConstant.commentset:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				if(excelCell == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					excelCell.setMemo((String)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.mergedelete:
			sheet.MergedRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		case OperatorConstant.frame:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.alignlevel:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.alignvertical:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
//		case OperatorConstant.rowsinsert:
//			break;
//		case OperatorConstant.rowsdelete:
//			break;
//		case OperatorConstant.colsinsert:
//			break;
//		case OperatorConstant.colsdelete:
//			break;
		case OperatorConstant.paste:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.copy:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.cut:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
//		case OperatorConstant.frozen:
//			break;
//		case OperatorConstant.unFrozen:
//			break;
//		case OperatorConstant.colswidth:
//			break;
//		case OperatorConstant.colshide:
//			break;	
//		case OperatorConstant.rowshide:
//			break;	
//		case OperatorConstant.colhideCancel:
//			break;	
//		case OperatorConstant.rowhideCancel:
//			break;	
//		case OperatorConstant.rowsheight:
//			break;
//		case OperatorConstant.addRowLine:
//			break;
//		case OperatorConstant.addColLine:
//			break;	
//		case OperatorConstant.colorset:
//		break;
		default:
			break;
		}
	}
	
	public void redo(VersionHistory versionHistory,int step,ExcelSheet sheet){
		int version = versionHistory.getVersion().get(step-1);
		History hisory= versionHistory.getMap().get(version+1);
		int operatorType = hisory.getOperatorType();
		List<ChangeArea> changeAreaList = hisory.getChangeAreaList();
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		versionHistory.getVersion().put(step, version+1);
		switch (operatorType) {
		case OperatorConstant.textData:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.fontsize:
		case OperatorConstant.fontfamily:
		case OperatorConstant.fontweight:
		case OperatorConstant.fontitalic:
		case OperatorConstant.fontcolor:
		case OperatorConstant.wordWrap:
		case OperatorConstant.alignlevel:
		case OperatorConstant.alignvertical:
		case OperatorConstant.fillbgcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				if(excelCell == null){
					excelCell = new ExcelCell();
					rowList.get(rowIndex).getCells().set(colIndex, excelCell);
				}
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.textDataformat:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;

		case OperatorConstant.commentset:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setMemo(changeArea.getUpdateValue().toString());
			}
			break;
		case OperatorConstant.merge:
			sheet.MergedRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		case OperatorConstant.mergedelete:
			sheet.SplitRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		
//		case OperatorConstant.rowsinsert:
//			break;
//		case OperatorConstant.rowsdelete:
//			break;
//		case OperatorConstant.colsinsert:
//			break;
//		case OperatorConstant.colsdelete:
//			break;
		case OperatorConstant.paste:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.copy:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.cut:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
//		case OperatorConstant.frozen:
//			break;
//		case OperatorConstant.unFrozen:
//			break;
//		case OperatorConstant.colswidth:
//			break;
//		case OperatorConstant.colshide:
//			break;	
//		case OperatorConstant.rowshide:
//			break;	
//		case OperatorConstant.colhideCancel:
//			break;	
//		case OperatorConstant.rowhideCancel:
//			break;	
//		case OperatorConstant.rowsheight:
//			break;
//		case OperatorConstant.addRowLine:
//			break;
//		case OperatorConstant.addColLine:
//			break;	
//		case OperatorConstant.colorset:
//		break;
		default:
			break;
		}
	}
	/**
	 * sheet保护
	 * @param protect
	 * @param excelBook
	 */
	public void protect(Protect protect,ExcelBook excelBook){
		excelBook.getSheets().get(0).setProtect(protect.isProtect());
		excelBook.getSheets().get(0).setPassword(protect.getPassword());
	}
	public void dataValidate(AreaSet areaSet, ExcelBook excelBook,String excelId) {
		List<Coordinate> coordinateList = areaSet.getCoordinate();
		ExcelSheet sheet = excelBook.getSheets().get(0);
		Data data = MemoryUtil.getDataValidateMap().get(excelId);
		if(data == null){
			data = new Data();
			MemoryUtil.getDataValidateMap().put(excelId, data);
		}
		Rule rule = new Rule();
		rule.setValidationType(areaSet.getRule().getValidationType());
		int index = data.getRuleList().indexOf(rule);
		if(index == -1){
			data.getRuleList().add(rule);
			index = data.getRuleList().size() -1 ;
		}
		String formal1 = areaSet.getRule().getFormula1();
		String formal2 = areaSet.getRule().getFormula2();
		if (areaSet.getRule().getValidationType() == 3) {
			formal1 = DataValidateUtil.display2Alias(formal1, sheet);
			if(formal1.contains("R") || formal1.contains("C")){
				setAffectRule(formal1, sheet, index);
			}
			formal2 = null;
		}
		rule.setFormula1(formal1);
		rule.setFormula2(formal2);
		for(Coordinate coordinate : coordinateList){
			int colStartIndex = coordinate.getStartCol();
			int rowStartIndex = coordinate.getStartRow();
			int colEndIndex = coordinate.getEndCol();
			int rowEndIndex = coordinate.getEndRow();
			
			ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
			ListHashMap<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
			for (int i = rowList.size() - 1; i <= rowEndIndex; i++) {
				sheet.addRow();
			}
			for (int i = columnList.size() - 1; i < colEndIndex; i++) {
				sheet.addColumn();
			}
			//整列
			if (rowEndIndex == -1) {
				for (int i = colStartIndex; i <= colEndIndex; i++) {
					Map<String, Integer> colMap = data.getColMap();
					String colCode = columnList.get(i).getCode();
					colMap.put(colCode, index);
					Map<String, Integer> rowMap = data.getRowMap();
					Set<String> keys = rowMap.keySet();
					for(String rowCode : keys){
						setCellMap(data, rowCode, colCode, index);
					}
				}
			}
			//整行
			else if (colEndIndex == -1 ) {
				for (int i = rowStartIndex; i <= rowEndIndex; i++) {
					Map<String, Integer> rowMap = data.getRowMap();
					String rowCode = rowList.get(i).getCode();
					rowMap.put(rowCode, index);
					Map<String, Integer> colMap = data.getColMap();
					Set<String> keys = colMap.keySet();
					for(String colCode : keys){
						setCellMap(data, rowCode, colCode, index);
					}
				}
			} else {
				//单元格
				for (int i = rowStartIndex; i <= rowEndIndex; i++) {
					ExcelRow excelRow = rowList.get(i);
					for (int j = colStartIndex; j <= colEndIndex; j++) {
						String rowCode = excelRow.getCode();
						String colCode = columnList.get(j).getCode();
						setCellMap(data, rowCode, colCode, index);
					}
				}
			}
		}
	}
	private void setAffectRule(String formal1,ExcelSheet excelSheet,int rule){
		String[] position = formal1.split(":");
		String start = position[0];
		String end = position[1];
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		if (start.contains("R") && start.contains("C") && end.contains("R") && end.contains("C")) {
			String aliasColStart = start.substring(1, start.lastIndexOf("R"));
			String aliasRowStart = start.substring(start.lastIndexOf("R")+1);
			String aliasColEnd = end.substring(1, end.lastIndexOf("R"));
			String aliasRowEnd = end.substring(end.lastIndexOf("R")+1);
			if(aliasColStart.equals(aliasColEnd)){
				setExps(rowList.get(rowList.getMaps().get(aliasRowStart)).getExps(), rule);
				setExps(rowList.get(rowList.getMaps().get(aliasRowEnd)).getExps(), rule);
			}else if(aliasRowStart.equals(aliasRowEnd)){
				setExps(colList.get(colList.getMaps().get(aliasColStart)).getExps(), rule);
				setExps(colList.get(colList.getMaps().get(aliasColEnd)).getExps(), rule);
			}
		}
	}

	private void setExps(Map<String, String> map, int rule) {
		String ruleList = map.get("rule");
		if (ruleList == null) {
			ruleList = "";
		}
		ruleList += rule + ",";
		map.put("rule", ruleList);
	}
	/**
	 * 处理单元格验证
	 * @param data
	 * @param rowCode
	 * @param colCode
	 * @param index
	 */
	
	private void setCellMap(Data data,String rowCode,String colCode,int index){
		Map<String, Integer> colRuleMap = data.getCellMap().get(rowCode);
		if (colRuleMap == null) {
			colRuleMap = new HashMap<String, Integer>();
			data.getCellMap().put(rowCode, colRuleMap);
		}
		colRuleMap.put(colCode, index);
	}
	
	
	public List<String> findSeq(String formal,ExcelSheet excelSheet) {
		if(formal == null || "".equals(formal)) return new ArrayList<String>();
		String[] position = formal.split(":");
		String start = position[0];
		String end = position[1];
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		List<String> seqList = new ArrayList<String>();
		if (start.contains("R") && start.contains("C") && end.contains("R") && end.contains("C")) {
			String aliasColStart = start.substring(1, start.lastIndexOf("R"));
			String aliasRowStart = start.substring(start.lastIndexOf("R")+1);
			int colStart = colList.getMaps().get(aliasColStart);
			int rowStart = rowList.getMaps().get(aliasRowStart);

			String aliasColEnd = end.substring(1, end.lastIndexOf("R"));
			String aliasRowEnd = end.substring(end.lastIndexOf("R")+1);
			int colEnd = colList.getMaps().get(aliasColEnd);
			int rowEnd = rowList.getMaps().get(aliasRowEnd);
			for (int i = rowStart; i <= rowEnd; i++) {
				ExcelRow excelRow = rowList.get(i);
				if(excelRow == null) break;
				List<ExcelCell> cellList = excelRow.getCells();
				for (int j = colStart; j <= colEnd; j++) {
					ExcelCell excelCell = cellList.get(j);
					if(excelCell == null) continue;
					seqList.add(excelCell.getText());
				}
			}
		}else {
			String aliasCode = start.substring(1);
			if(start.contains("R")){ 
				int row = rowList.getMaps().get(aliasCode);
				ExcelRow excelRow = rowList.get(row);
				for(int i = 0;i<excelRow.getCells().size();i++){
					ExcelCell excelCell = excelRow.getCells().get(i);
					if(excelCell == null) continue;
					seqList.add(excelCell.getText());
				}
			}else{
				int col = colList.getMaps().get(aliasCode);
				for(int i = 0; i<excelSheet.getMaxrow();i++){
					ExcelCell excelCell = rowList.get(i).getCells().get(col);
					if(excelCell == null) continue;
					seqList.add(excelCell.getText());
				}
			}
		}
		return seqList;
	}
	public int getRule(Data data,Frozen frozen,ExcelSheet excelSheet){
		int col = frozen.getOprCol();
		int row = frozen.getOprRow();
		if(col == -1){
			return data.getRowMap().get(excelSheet.getRows().get(row).getCode());
		}else if(row == -1){
			return data.getColMap().get(excelSheet.getCols().get(col).getCode());
		}else{
			Map<String, Map<String, Integer>> cellMap = data.getCellMap();
			Map<String, Integer> colRuleMap = cellMap.get(excelSheet.getRows().get(row).getCode());
			if(colRuleMap == null){
				Integer rRule = data.getRowMap().get(excelSheet.getRows().get(row).getCode());
				return (rRule == null) ? data.getColMap().get(excelSheet.getCols().get(col).getCode()) : rRule;
			}else{
				return colRuleMap.get(excelSheet.getCols().get(col).getCode());
			}
		}
	}
	
	
}
