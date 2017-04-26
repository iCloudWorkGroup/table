package com.acmr.mq.consumer.queue;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.model.AddLine;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.CellService;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.SheetService;
import com.acmr.mq.AffectCell;
import com.acmr.mq.Model;
import com.alibaba.fastjson.JSON;

public class WorkerThread2 implements Runnable{
	private static Logger logger = Logger.getLogger(QueueReceiver.class);
	private int step;  
	private MemcachedClient memcachedClient;
	private String key;
	private HandleExcelService handleExcelService;
	private CellService cellService;
	private PasteService pasteService;
	private SheetService sheetService;
	private Model model;
	
	
    
    public WorkerThread2(int step,MemcachedClient memcachedClient,String key,HandleExcelService handleExcelService,
    		CellService cellService,PasteService pasteService,SheetService sheetService,Model model){  
        this.step=step;
        this.memcachedClient = memcachedClient;
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
    		int memStep = (Integer) memcachedClient.get(key);
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
		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
		VersionHistory versionHistory = (VersionHistory) memcachedClient.get(excelId+"_history");
		Cell cell = null;
		switch (reqPath) {
		case OperatorConstant.textData:
			cell = (Cell) model.getObject();
			handleExcelService.data(cell, excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.fontsize:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_size, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.fontfamily:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_family, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.fontweight:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_weight, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.fontitalic:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_italic, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.fontcolor:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.font_color, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.wordWrap:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.word_wrap, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;

		case OperatorConstant.fillbgcolor:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.textDataformat:
			CellFormate cellFormate = (CellFormate) model.getObject();
			handleExcelService.setCellFormate(cellFormate, excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;

		case OperatorConstant.commentset:
			Comment comment = (Comment) model.getObject();
			handleExcelService.setComment(excelBook, comment,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.merge:
			cell = (Cell) model.getObject();
			cellService.mergeCell(excelBook.getSheets().get(0), cell,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.mergedelete:
			cell = (Cell) model.getObject();
			cellService.splitCell(excelBook.getSheets().get(0), cell,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.frame:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.frame, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.alignlevel:
			cell = (Cell) model.getObject();
			handleExcelService.updateCells(CellUpdateType.align_level, cell,excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
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
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.copy:
			Copy copy = (Copy) model.getObject();
			pasteService.copy(copy, excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.cut:
			Copy copy2 = (Copy) model.getObject();
			pasteService.cut(copy2, excelBook,versionHistory,step);
			memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
			break;
		case OperatorConstant.frozen:
			Frozen frozen = (Frozen) model.getObject();
			sheetService.frozen(excelBook.getSheets().get(0), frozen);
			break;
		case OperatorConstant.unFrozen:
			excelBook.getSheets().get(0).setFreeze(null);
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
		    memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
		break;
		case OperatorConstant.redo:
		    sheetService.redo(versionHistory, step, excelBook.getSheets().get(0));
		    memcachedClient.set(excelId+"_history", Constant.MEMCACHED_EXP_TIME, versionHistory);
		break;
		default:
			break;
		}
		//System.out.println(JSON.toJSONString(versionHistory));
		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME, excelBook);
		System.out.println(step + "结束执行");
		logger.info("**********end excelId : "+excelId + " === step : " + step + "== reqPath : "+ reqPath);
		memcachedClient.set(excelId + "_ope", Constant.MEMCACHED_EXP_TIME, step);
	}
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
}
