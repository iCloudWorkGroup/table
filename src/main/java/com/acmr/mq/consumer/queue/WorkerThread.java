package com.acmr.mq.consumer.queue;

import net.spy.memcached.MemcachedClient;
import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.service.CellService;
import com.acmr.excel.service.HandleExcelService;
import com.acmr.excel.service.HandleExcelService.CellUpdateType;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.SheetService;
import com.acmr.mq.AffectCell;
import com.acmr.mq.Model;

public class WorkerThread implements Runnable{

	private int step;  
	private MemcachedClient memcachedClient;
	private String key;
	private HandleExcelService handleExcelService;
	private CellService cellService;
	private PasteService pasteService;
	private SheetService sheetService;
	private Model model;
	
	
    
    public WorkerThread(int step,MemcachedClient memcachedClient,String key,HandleExcelService handleExcelService,
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
//		int reqPath = model.getReqPath();
//		String excelId = model.getExcelId();
//		int step = model.getStep();
//		ExcelBook excelBook = (ExcelBook) memcachedClient.get(excelId);
//		Cell cell = null;
//		AffectCell affectCell = new AffectCell();
//		switch (reqPath) {
//		case "/text/data":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getStartX());
//			handleExcelService.data(cell, excelBook);
//			break;
//		case "/text/fontsize":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.font_size, cell,excelBook);
//			break;
//		case "/text/fontfamily":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.font_family, cell,excelBook);
//			break;
//		case "/text/fontweight":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.font_weight, cell,excelBook);
//			break;
//		case "/text/fontitalic":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.font_italic, cell,excelBook);
//			break;
//		case "/text/fontcolor":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.font_color, cell,excelBook);
//			break;
//		case "/text/wordWrap":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.word_wrap, cell,excelBook);
//			break;
//
//		case "/text/fillbgcolor":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.fill_bgcolor, cell,excelBook);
//			break;
//		case "/text/dataformat":
//			CellFormate cellFormate = (CellFormate) model.getObject();
//			affectCell.setStartRowAlias(cellFormate.getCoordinate().getStartRowAlais());
//			affectCell.setEndRowAlias(cellFormate.getCoordinate().getEndRowAlais());
//			affectCell.setStartColAlias(cellFormate.getCoordinate().getStartColAlais());
//			affectCell.setEndColAlias(cellFormate.getCoordinate().getEndColAlais());
//			handleExcelService.setCellFormate(cellFormate, excelBook);
//			break;
//
//		case "/text/commentset":
//			Comment comment = (Comment) model.getObject();
//			affectCell.setStartRowAlias(comment.getCoordinate().getStartRowAlais());
//			affectCell.setEndRowAlias(comment.getCoordinate().getEndRowAlais());
//			affectCell.setStartColAlias(comment.getCoordinate().getStartColAlais());
//			affectCell.setEndColAlias(comment.getCoordinate().getEndColAlais());
//			handleExcelService.setComment(excelBook, comment);
//			break;
//		case "/cell/merge":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			cellService.mergeCell(excelBook.getSheets().get(0), cell);
//			break;
//		case "/cell/mergedelete":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			cellService.splitCell(excelBook.getSheets().get(0), cell);
//			break;
//		case "/cell/frame":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.frame, cell,excelBook);
//			break;
//		case "/cell/alignlevel":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.align_level, cell,excelBook);
//			break;
//		case "/cell/alignvertical":
//			cell = (Cell) model.getObject();
//			affectCell.setStartRowAlias(cell.getCoordinate().getStartY());
//			affectCell.setEndRowAlias(cell.getCoordinate().getEndY());
//			affectCell.setStartColAlias(cell.getCoordinate().getStartX());
//			affectCell.setEndColAlias(cell.getCoordinate().getEndX());
//			handleExcelService.updateCells(CellUpdateType.align_vertical, cell,excelBook);
//			break;
//		case "/cell/rowsinsert":
//			RowOperate rowOperate = (RowOperate) model.getObject();
//			affectCell.setStartRowAlias(rowOperate.getRowAlias());
////			affectCell.setEndRowAlias(rowOperate.getRowAlias());
////			affectCell.setStartColAlias("MIN");
////			affectCell.setEndColAlias("MAX");
//			affectCell.setType("rows_insert");
//			cellService.addRow(excelBook.getSheets().get(0), rowOperate);
//			break;
//		case "/cell/rowsdelete":
//			RowOperate rowOperate2 = (RowOperate) model.getObject();
//			affectCell.setStartRowAlias(rowOperate2.getRowAlias());
////			affectCell.setEndRowAlias(rowOperate2.getRowAlias());
////			affectCell.setStartColAlias("MIN");
////			affectCell.setEndColAlias("MAX");
//			affectCell.setType("rows_delete");
//			cellService.deleteRow(excelBook.getSheets().get(0), rowOperate2);
//			break;
//		case "/cell/colsinsert":
//			ColOperate colOperate = (ColOperate) model.getObject();
////			affectCell.setStartRowAlias("MIN");
////			affectCell.setEndRowAlias("MAX");
//			affectCell.setStartColAlias(colOperate.getColAlias());
//			//affectCell.setEndColAlias(colOperate.getColAlias());
//			affectCell.setType("cols_insert");
//			cellService.addCol(excelBook.getSheets().get(0), colOperate);
//			break;
//		case "/cell/colsdelete":
//			ColOperate colOperate2 = (ColOperate) model.getObject();
////			affectCell.setStartRowAlias("MIN");
////			affectCell.setEndRowAlias("MAX");
//			affectCell.setStartColAlias(colOperate2.getColAlias());
//			//affectCell.setEndColAlias(colOperate2.getColAlias());
//			affectCell.setType("cols_delete");
//			cellService.deleteCol(excelBook.getSheets().get(0), colOperate2);
//			break;
//		case "/plate/paste":
//			Paste paste = (Paste) model.getObject();
//			affectCell.setStartRowAlias(paste.getStartRowAlias());
//			affectCell.setStartColAlias(paste.getStartColAlias());
//			pasteService.data(paste, excelBook);
//			break;
//		case "/plate/copy":
//			Copy copy = (Copy) model.getObject();
//			affectCell.setStartRowAlias(copy.getTarget().getRowAlias());
//			affectCell.setStartColAlias(copy.getTarget().getColAlias());
//			pasteService.copy(copy, excelBook);
//			break;
//		case "/plate/cut":
//			Copy copy2 = (Copy) model.getObject();
//			affectCell.setStartRowAlias(copy2.getTarget().getRowAlias());
//			affectCell.setStartColAlias(copy2.getTarget().getColAlias());
//			pasteService.cut(copy2, excelBook);
//			break;
//		case "/sheet/frozen":
//			Frozen frozen = (Frozen) model.getObject();
//			affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(frozen.getFrozenX());
//			sheetService.frozen(excelBook.getSheets().get(0), frozen.getFrozenY(), frozen.getFrozenX(), frozen.getFrozenY() ,frozen.getStartX());
//			break;
//		case "/sheet/unFrozen":
//			Frozen frozen2 = (Frozen) model.getObject();
//			affectCell.setStartRowAlias(frozen2.getFrozenY());
//			affectCell.setStartColAlias(frozen2.getFrozenX());
//			excelBook.getSheets().get(0).setFreeze(null);
//			break;
//		case "/cell/colswidth":
//			ColWidth colWidth = (ColWidth) model.getObject();
//			//affectCell.setStartRowAlias(frozen.getFrozenY());
//			affectCell.setStartColAlias(colWidth.getColAlias());
//			affectCell.setType("cols_width");
//			cellService.controlColWidth(excelBook.getSheets().get(0), colWidth.getColAlias(), colWidth.getOffset());
//			break;
//		case "/cell/rowsheight":
//			RowHeight rowHeight = (RowHeight) model.getObject();
//			affectCell.setStartRowAlias(rowHeight.getRowAlias());
//			affectCell.setType("rows_height");
//			cellService.controlRowHeight(excelBook.getSheets().get(0), rowHeight.getRowAlias(), rowHeight.getOffset());
//			break;
//		case "/sheet/addRowLine":
//			RowLine rowLine = (RowLine) model.getObject();
//			affectCell.setStartRowAlias(rowLine.getRowNum());
//			//affectCell.setStartColAlias(colWidth.getColAlias());
//			String rowNum = rowLine.getRowNum();
//			int rn = Integer.valueOf(rowNum);
//			sheetService.addRowLine(excelBook.getSheets().get(0),rn);
//			break;
//		case "disablePaste":
//			//Model mod =  (Model)model.getObject();
//			break;
//		default:
//			break;
//		}
//		memcachedClient.set(excelId, Constant.MEMCACHED_EXP_TIME, excelBook);
//		//memcachedClient.set(excelId+"_"+step, Constant.MEMCACHED_EXP_TIME, affectCell);
//		//System.out.println(JSON.toJSONString(excelBook));
//		System.out.println(step + "结束执行");
//		memcachedClient.set(excelId + "_ope", Constant.MEMCACHED_EXP_TIME, step);
	}
}
