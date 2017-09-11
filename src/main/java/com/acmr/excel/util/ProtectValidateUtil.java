package com.acmr.excel.util;

import java.util.List;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import com.acmr.mq.Model;

public class ProtectValidateUtil {
	
	
	
	
	/**
     * 校验复制粘贴和剪切粘贴
     * @param model
     * @param rowList
     * @return
     */
    public static boolean validateCopyOrCut(Model model,List<ExcelRow> rowList){
    	Copy copy = (Copy) model.getObject();
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int targetRowIndex = copy.getTarget().getOprRow();
		int targetColIndex = copy.getTarget().getOprCol();
		int targetRLen = endRowIndex - startRowIndex;
		int targetCLen = endColIndex - startColIndex;
		if (model.getReqPath() == OperatorConstant.cut) {
			if (rowList.size() <= startRowIndex)
				return false;
			for (int i = startRowIndex; i <= endRowIndex; i++) {
				List<ExcelCell> cellList = rowList.get(i).getCells();
				if (cellList.size() <= startColIndex)
					return false;
				for (int j = startColIndex; j <= endColIndex; j++) {
					if (cellList.get(j).getCellstyle().isLocked())
						return false;
				}
			}
		}
		if (rowList.size() <= targetRowIndex)
			return false;
		for (int i = targetRowIndex; i <= targetRowIndex+targetRLen; i++) {
			List<ExcelCell> cellList = rowList.get(i).getCells();
			if (rowList.size() <= targetColIndex)
				return false;
			for (int j = targetColIndex; j <= targetColIndex+targetCLen ;j++) {
				if (cellList.get(j).getCellstyle().isLocked())
					return false;
			}
		}
		return true;
    }
    /**
     * 校验外部粘贴
     * @param model
     * @param rowList
     * @return
     */
	public static boolean validatePaste(Model model, List<ExcelRow> rowList) {
		Paste paste = (Paste) model.getObject();
		int oprr = paste.getOprRow();
		int oprc = paste.getOprCol();
		if (rowList.size() <= oprr)
			return false;
		for (int i = oprr; i < oprr + paste.getRowLen(); i++) {
			List<ExcelCell> cellList = rowList.get(i).getCells();
			if (cellList.size() <= oprc)
				return false;
			for (int j = oprc; j < oprc + paste.getColLen(); j++) {
				if (cellList.get(j) == null) {
					return false;
				}
				if (cellList.get(j).getCellstyle().isLocked())
					return false;
			}
		}
		return true;
	}
    /**
     * 校验行
     * @param model
     * @param rowList
     * @return
     */
    public static boolean validateRow(int row,List<ExcelRow> rowList){
    	if (rowList.size() <= row)
			return false;
//		if (rowList.get(row) == null)
//			return false;
		if(rowList.get(row).getCellstyle().isLocked())
			return false;
    	return true;
    }
    
    /**
     * 校验列
     * @param model
     * @param rowList
     * @return
     */
    public static boolean validateCol(int col,List<ExcelColumn> colList){
    	if (colList.size() <= col)
			return false;
//		if (colList.get(col) == null)
//			return false;
		if(colList.get(col).getCellstyle().isLocked())
			return false;
    	return true;
    }
    
    /**
     * 校验操作
     * @param colStartIndex
     * @param rowStartIndex
     * @param colEndIndex
     * @param rowEndIndex
     * @param sheet
     * @return
     */
    public static boolean validateOpr(int colStartIndex, int rowStartIndex, int colEndIndex, int rowEndIndex,ExcelSheet sheet) {
		List<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		List<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
		//整列
		if (rowEndIndex == -1) {
			rowEndIndex = rowList.size() - 1;
			for (int i = colStartIndex; i <= colEndIndex; i++) {
				if (columnList.get(i).getCellstyle().isLocked())
					return false;
			}
		}
		boolean rowFlag = false;
		if (colEndIndex == -1 ) {
			colEndIndex = columnList.size() - 1;
			rowFlag = true;
		} 
		for(int i = rowStartIndex ; i <= rowEndIndex ;i++){
			ExcelRow excelRow = rowList.get(i);
			//整行
			if(rowFlag){
				if (rowList.get(i).getCellstyle().isLocked())
					return false;
			}
			List<ExcelCell> cellList = excelRow.getCells();
//			if(cellList.size() == 0)
//				return false;
			for(int j = colStartIndex;j<=colEndIndex;j++){
				ExcelCell excelCell = cellList.get(j);
				if(excelCell == null) return false;
				if ( excelCell.getCellstyle().isLocked()) 
					return false;
			}
		}
		return true;
    }
    
    /**
     * 校验受保护的状态
     * @param sheet
     * @param type
     * @return
     */
    public static boolean validateStatus(ExcelSheet sheet,int type){
    	if (sheet.isProtect()
				&& (type == OperatorConstant.textData
						|| type == OperatorConstant.paste
						|| type == OperatorConstant.copy || type == OperatorConstant.cut)) 
			return true;
    	return false;
    }
    
    
    
}
