package com.acmr.mq.consumer.queue;


import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.AddLine;
import com.acmr.excel.model.AreaSet;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.Protect;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.CellService;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.SheetService;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.ProtectValidateUtil;
import com.acmr.mq.Model;

public class WorkerThread2 implements Runnable{
	private static Logger logger = Logger.getLogger(QueueReceiver.class);
	private int step;  
	private StoreService storeService;
	private String key;
	private HandleExcelService handleExcelService;
	private CellService cellService;
	private PasteService pasteService;
	private SheetService sheetService;
	private Model model;
	
	
    
    public WorkerThread2(int step,StoreService storeService,String key,HandleExcelService handleExcelService,
    		CellService cellService,PasteService pasteService,SheetService sheetService,Model model){  
        this.step=step;
        this.storeService = storeService;
        this.key = key;
        this.handleExcelService = handleExcelService;
        this.cellService = cellService;
        this.pasteService = pasteService;
        this.sheetService = sheetService;
        this.model = model;
    }  
   
    @Override  
    public void run() {  
    	while(true){
    		int memStep = (Integer) storeService.get(key);
    		if(memStep + 1 == step){
    			System.out.println(step + "开始执行");
    			logger.info("**********begin excelId : "+model.getExcelId() + " === step : " + step + "== reqPath : "+ model.getReqPath());
    			handleMessage(model);
    			return;
    		}else{
    			processCommand(10);
    			continue;
    		}
    	}
    }  
   
    private void processCommand(int n) {  
        try {  
            Thread.sleep(n);  
        } catch (InterruptedException e) {  
            e.printStackTrace();  
        }  
    }  
   
    @Override  
    public String toString(){  
        return this.step+"";  
    }
    private void handleMessage(Model model) {
		int reqPath = model.getReqPath();
		String excelId = model.getExcelId();
		int step = model.getStep();
		ExcelBook excelBook = (ExcelBook) storeService.get(excelId);
		if (!validate(model, excelBook))
			return;
		VersionHistory versionHistory = (VersionHistory) storeService.get(excelId+"_history");
		Cell cell = null;
		switch (reqPath) {
		case OperatorConstant.textData:
			cell = (Cell) model.getObject();
			handleExcelService.data(cell, excelBook,versionHistory,step);
			storeService.set(excelId+"_history",versionHistory);
			break;
		case OperatorConstant.fontsize:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_size, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history",  versionHistory);
			break;
		case OperatorConstant.fontfamily:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_family, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.fontweight: 
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_weight, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.fontitalic:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_italic, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.UNDERLINE:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_underline, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;	
		case OperatorConstant.fontcolor:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_color, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.wordWrap:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.word_wrap, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;

		case OperatorConstant.fillbgcolor:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.textDataformat:
			CellFormate cellFormate = (CellFormate) model.getObject();
			handleExcelService.setCellFormate(cellFormate, excelBook,versionHistory,step);
			storeService.set(excelId+"_history",versionHistory);
			break;

		case OperatorConstant.commentset:
			Comment comment = (Comment) model.getObject();
			handleExcelService.setComment(excelBook, comment,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.merge:
			cell = (Cell) model.getObject();
			cellService.mergeCell(excelBook.getSheets().get(0), cell,versionHistory,step);
			storeService.set(excelId+"_history",versionHistory);
			break;
		case OperatorConstant.mergedelete:
			cell = (Cell) model.getObject();
			cellService.splitCell(excelBook.getSheets().get(0), cell,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.frame:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.frame, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history",  versionHistory);
			break;
		case OperatorConstant.alignlevel:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.align_level, cell,excelBook,versionHistory,step);
			storeService.set(excelId+"_history",  versionHistory);
			break;
		case OperatorConstant.alignvertical:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.align_vertical, cell,excelBook,versionHistory,step);
			break;
		case OperatorConstant.rowsinsert:
			RowOperate rowOperate = (RowOperate) model.getObject();
			cellService.addRow(excelBook.getSheets().get(0), rowOperate);
			break;
		case OperatorConstant.rowsdelete:
			RowOperate rowOperate2 = (RowOperate) model.getObject();
			cellService.deleteRow(excelBook.getSheets().get(0), rowOperate2);
			break;
		case OperatorConstant.colsinsert:
			ColOperate colOperate = (ColOperate) model.getObject();
			cellService.addCol(excelBook.getSheets().get(0), colOperate);
			break;
		case OperatorConstant.colsdelete:
			ColOperate colOperate2 = (ColOperate) model.getObject();
			cellService.deleteCol(excelBook.getSheets().get(0), colOperate2);
			break;
		case OperatorConstant.paste:
			Paste paste = (Paste) model.getObject();
			pasteService.data(paste, excelBook,versionHistory,step);
			storeService.set(excelId+"_history",  versionHistory);
			break;
		case OperatorConstant.copy:
			Copy copy = (Copy) model.getObject();
			pasteService.copy(copy, excelBook,versionHistory,step);
			storeService.set(excelId+"_history", versionHistory);
			break;
		case OperatorConstant.cut:
			Copy copy2 = (Copy) model.getObject();
			pasteService.cut(copy2, excelBook,versionHistory,step);
			storeService.set(excelId+"_history",versionHistory);
			break;
		case OperatorConstant.frozen:
			Frozen frozen = (Frozen) model.getObject();
			sheetService.frozen(excelBook.getSheets().get(0), frozen);
			break;
		case OperatorConstant.unFrozen:
			excelBook.getSheets().get(0).setFreeze(null);
			Map<String, String> map = excelBook.getSheets().get(0).getExps();
			map.remove("fr");
			map.remove("fc");
			break;
		case OperatorConstant.colswidth:
			ColWidth colWidth = (ColWidth) model.getObject();
			cellService.controlColWidth(excelBook.getSheets().get(0), colWidth);
			break;
		case OperatorConstant.colshide:
			ColOperate colHide = (ColOperate) model.getObject();
			cellService.colHide(excelBook.getSheets().get(0), colHide);
			break;	
		case OperatorConstant.rowshide:
			RowOperate rowHide = (RowOperate) model.getObject();
			cellService.rowHide(excelBook.getSheets().get(0), rowHide);
			break;	
		case OperatorConstant.colhideCancel:
			ColOperate colhideCancel = (ColOperate) model.getObject();
			sheetService.cancelColHide(excelBook.getSheets().get(0), colhideCancel);
			break;	
		case OperatorConstant.rowhideCancel:
			RowOperate rowhideCancel = (RowOperate) model.getObject();
			sheetService.cancelRowHide(excelBook.getSheets().get(0), rowhideCancel);
			break;	
		case OperatorConstant.rowsheight:
			RowHeight rowHeight = (RowHeight) model.getObject();
			cellService.controlRowHeight(excelBook.getSheets().get(0), rowHeight);
			break;
		case OperatorConstant.addRowLine:
			RowLine rowLine = (RowLine) model.getObject();
			int rowNum = rowLine.getNum();
			sheetService.addRowLine(excelBook.getSheets().get(0),rowNum);
			break;
		case OperatorConstant.addColLine:
			AddLine addLine = (AddLine) model.getObject();
			int num = addLine.getNum();
			sheetService.addColLine(excelBook.getSheets().get(0), num);
			break;	
		case OperatorConstant.colorset:
			cell = (Cell) model.getObject();
			handleExcelService.colorSet(cell, excelBook);	
		case OperatorConstant.undo:
		    sheetService.undo(versionHistory, step, excelBook.getSheets().get(0));
		    storeService.set(excelId+"_history", versionHistory);
		break;
		case OperatorConstant.redo:
		    sheetService.redo(versionHistory, step, excelBook.getSheets().get(0));
		    storeService.set(excelId+"_history",  versionHistory);
		break;
		case OperatorConstant.batchcolorset:
			AreaSet colorSet= (AreaSet) model.getObject();
			handleExcelService.areaSet(colorSet, excelBook,OperatorConstant.batchcolorset);
		break;	
		case OperatorConstant.CLEANDATA:
			AreaSet areaDel= (AreaSet) model.getObject();
			handleExcelService.areaSet(areaDel, excelBook,OperatorConstant.CLEANDATA);
		break;
		case OperatorConstant.CELLLOCK:
			AreaSet cellLock= (AreaSet) model.getObject();
			handleExcelService.areaSet(cellLock, excelBook,OperatorConstant.CELLLOCK);
		break;
		case OperatorConstant.PROTECT:
			Protect protect= (Protect) model.getObject();
			sheetService.protect(protect, excelBook);
		break;
		default:
			break;
		}
		//System.out.println(JSON.toJSONString(versionHistory));
		storeService.set(excelId,excelBook);
		System.out.println(step + "结束执行");
		logger.info("**********end excelId : "+excelId + " === step : " + step + "== reqPath : "+ reqPath);
		storeService.set(excelId + "_ope",  step);
	}
    
    
//    private boolean validate(Model model,ExcelBook excelBook){
//    	int reqPath = model.getReqPath();
//		ExcelSheet sheet = excelBook.getSheets().get(0);
//		if(!sheet.isProtect()){
//			return true;
//		}
//		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
//		ListHashMap<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
//		switch (reqPath) {
//		case OperatorConstant.textData:
//		case OperatorConstant.fontsize:
//		case OperatorConstant.fontfamily:
//		case OperatorConstant.fontweight:
//		case OperatorConstant.fontitalic:
//		case OperatorConstant.UNDERLINE:
//		case OperatorConstant.fontcolor:
//		case OperatorConstant.wordWrap:
//		case OperatorConstant.fillbgcolor:
//		case OperatorConstant.merge:
//		case OperatorConstant.mergedelete:
//		case OperatorConstant.frame:
//		case OperatorConstant.alignlevel:
//		case OperatorConstant.alignvertical:
//			Cell cell = (Cell) model.getObject();
//			 int rowStartIndex = cell.getCoordinate().getStartRow();
//			 int colStartIndex = cell.getCoordinate().getStartCol();
//			 int rowEndIndex = cell.getCoordinate().getEndRow();
//			 int colEndIndex = cell.getCoordinate().getEndCol();
//			 if(!ProtectValidateUtil.validateOpr(colStartIndex, rowStartIndex, colEndIndex, rowEndIndex, sheet))
//				 return false;
//			break;
//		case OperatorConstant.textDataformat:
//			CellFormate cellFormate = (CellFormate) model.getObject();
//			rowStartIndex = cellFormate.getCoordinate().getStartRow();
//			rowEndIndex = cellFormate.getCoordinate().getEndRow();
//			colStartIndex = cellFormate.getCoordinate().getStartCol();
//			colEndIndex = cellFormate.getCoordinate().getEndCol();
//			 if(!ProtectValidateUtil.validateOpr(colStartIndex, rowStartIndex, colEndIndex, rowEndIndex, sheet))
//				 return false;
//			break;
//
//		case OperatorConstant.commentset:
//			Comment comment = (Comment) model.getObject();
//			rowStartIndex = comment.getCoordinate().getStartRow();
//			rowEndIndex = comment.getCoordinate().getEndRow();
//			colStartIndex = comment.getCoordinate().getStartCol();
//			colEndIndex = comment.getCoordinate().getEndCol();
//			 return ProtectValidateUtil.validateOpr(colStartIndex, rowStartIndex, colEndIndex, rowEndIndex, sheet);
//		case OperatorConstant.rowsinsert:
//		case OperatorConstant.rowsdelete:
//		case OperatorConstant.colsinsert:
//		case OperatorConstant.colsdelete:
//			return false;
//		case OperatorConstant.paste:
//			return ProtectValidateUtil.validatePaste(model, rowList);
//		case OperatorConstant.copy:
//		case OperatorConstant.cut:
//			return ProtectValidateUtil.validateCopyOrCut(model, rowList);
//		case OperatorConstant.frozen:
//		case OperatorConstant.unFrozen:
//			Frozen frozen = (Frozen) model.getObject();
//			if (!ProtectValidateUtil.validateRow(frozen.getOprRow(), rowList))
//				return false;
//			if (!ProtectValidateUtil.validateCol(frozen.getOprCol(), columnList))
//				return false;
//			break;
//		
//		case OperatorConstant.colswidth:
//			ColWidth colWidth = (ColWidth) model.getObject();
//			return ProtectValidateUtil.validateCol(colWidth.getCol(), columnList);
//		case OperatorConstant.colshide:
//		case OperatorConstant.colhideCancel:	
//			ColOperate colHide = (ColOperate) model.getObject();
//			return ProtectValidateUtil.validateCol(colHide.getCol(), columnList);
//		case OperatorConstant.rowshide:
//		case OperatorConstant.rowhideCancel:
//			RowOperate rowHide = (RowOperate) model.getObject();
//			return ProtectValidateUtil.validateRow(rowHide.getRow(), rowList);
//		case OperatorConstant.rowsheight:
//			RowHeight rowHeight = (RowHeight) model.getObject();
//			return ProtectValidateUtil.validateRow(rowHeight.getRow(), rowList);
//		case OperatorConstant.addRowLine:
//			return false;
//		case OperatorConstant.addColLine:
//			return false;
//		case OperatorConstant.undo:
//		case OperatorConstant.redo:
//		   return false;
//		case OperatorConstant.batchcolorset:
//		case OperatorConstant.CLEANDATA:
//			AreaSet areaSet= (AreaSet) model.getObject();
//			List<Coordinate> coordinateList = areaSet.getCoordinate();
//			for (Coordinate coordinate : coordinateList) {
//				colStartIndex = coordinate.getStartCol();
//				rowStartIndex = coordinate.getStartRow();
//				colEndIndex = coordinate.getEndCol();
//				rowEndIndex = coordinate.getEndRow();
//				if (!ProtectValidateUtil.validateOpr(colStartIndex, rowStartIndex, colEndIndex, rowEndIndex, sheet))
//					return false;
//			}
//		break;
//		default:
//			break;
//		}
//		return true;
//    }
    /**
     * 保护校验
     * @param model
     * @param excelBook
     * @return
     */
	private boolean validate(Model model, ExcelBook excelBook) {
		int reqPath = model.getReqPath();
		ExcelSheet sheet = excelBook.getSheets().get(0);
		if (!sheet.isProtect())
			return true;
		if (reqPath == OperatorConstant.PROTECT)
			return true;
		if (!ProtectValidateUtil.validateStatus(sheet, reqPath))
			return false;
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) sheet.getRows();
		switch (reqPath) {
		case OperatorConstant.textData:
			Cell cell = (Cell) model.getObject();
			int rowStartIndex = cell.getCoordinate().getStartRow();
			int colStartIndex = cell.getCoordinate().getStartCol();
			int rowEndIndex = cell.getCoordinate().getEndRow();
			int colEndIndex = cell.getCoordinate().getEndCol();
			return ProtectValidateUtil.validateOpr(colStartIndex, rowStartIndex, colEndIndex, rowEndIndex, sheet);
		case OperatorConstant.paste:
			return ProtectValidateUtil.validatePaste(model, rowList);
		case OperatorConstant.copy:
		case OperatorConstant.cut:
			return ProtectValidateUtil.validateCopyOrCut(model, rowList);
		default:
			break;
		}
		return true;
	}
    
    
}
