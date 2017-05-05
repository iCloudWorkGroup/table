package com.acmr.excel.controller;

import java.io.IOException;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.position.OpenExcel;
import com.acmr.excel.service.ExcelService;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.StringUtil;

/**
 * SHEET操作
 * 
 * @author jinhr
 *
 */
@Controller
@RequestMapping("/sheet")
public class SheetController extends BaseController {
	@Resource
	private StoreService storeService;
	@Resource
	private PasteService pasteService; 
	 @Resource
	private ExcelService excelService;

	/**
	 * 新建sheet
	 * 
	 * @throws IOException
	 */
	public void create(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 修改sheet
	 * 
	 */
	public void update(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 删除sheet
	 * 
	 * @throws IOException
	 */
	public void delete(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 冻结
	 * 
	 * @throws Exception
	 */
    @RequestMapping("/frozen")
	public void frozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.frozen);
	}

	/**
	 * 取消冻结
	 * 
	 * @throws Exception
	 */
    @RequestMapping("/unfrozen")
	public void unFrozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.unFrozen);
	}

	
	
	
	/**
	 * 回退
	 */
    @RequestMapping("/undo")
	public void undo(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		this.assembleData(req, resp,null,OperatorConstant.undo);
	}
	/**
	 * 前进
	 */
    @RequestMapping("/redo")
	public void redo(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		this.assembleData(req, resp,null,OperatorConstant.redo);
	}
	/**
	 * 外部粘贴
	 * @throws IOException
	 */
	@RequestMapping("/paste")
	public void paste(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		Paste paste = getJsonDataParameter(req, Paste.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(req.getHeader("excelId"));
		boolean isAblePasteResult = pasteService.isAblePaste(paste, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, paste, OperatorConstant.paste);
			this.sendJson(resp, true);
		}else{
			this.sendJson(resp, false);
		}
		
	}
	/**
	 * 内部复制粘贴
	 * @throws IOException
	 */
	@RequestMapping("/copy")
	public void copy(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(req.getHeader("excelId"));
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.copy);
			this.sendJson(resp, true);
		}else{
			this.sendJson(resp, false);
		}
		
	}
	/**
	 * 剪切粘贴
	 * @throws IOException
	 */
	@RequestMapping("/cut")
	public void cut(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(req.getHeader("excelId"));
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.cut);
			this.sendJson(resp, true);
		}else{
			this.sendJson(resp, false);
		}
	}
	
	
	/**
	 * 通过像素动态加载excel
	 */
	@RequestMapping("/area")
	public void openexcel(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		String excelId = req.getHeader("excelId");
		int memStep = (int)storeService.get(excelId+"_ope");
		String curStep = req.getHeader("step");
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		ExcelBook excelBook = (ExcelBook) storeService.get(excelId);
		JsonReturn data = new JsonReturn("");
		if (cStep == memStep) {
			if (excelBook != null) {
				ExcelSheet excelSheet = excelBook.getSheets().get(0);
				ReturnParam returnParam = new ReturnParam();
				CompleteExcel excel = new CompleteExcel();
				SpreadSheet spreadSheet = new SpreadSheet();
				excel.getSpreadSheet().add(spreadSheet);
				spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam);
				data.setReturncode(Constant.SUCCESS_CODE);
				data.setReturndata(excel);
				data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
				data.setMaxRowPixel(returnParam.getMaxRowPixel());
				storeService.set(excelId,excelBook);
			} else {
				data.setReturncode(Constant.CACHE_INVALID_CODE);
				data.setReturndata(Constant.CACHE_INVALID_MSG);
			}
		}else{
			for (int i = 0; i < 100; i++) {
				int mStep = (int)storeService.get(excelId+"_ope");
				if(cStep == mStep){
					if (excelBook != null) {
						ExcelSheet excelSheet = excelBook.getSheets().get(0);
						ReturnParam returnParam = new ReturnParam();
						CompleteExcel excel = new CompleteExcel();
						SpreadSheet spreadSheet = new SpreadSheet();
						excel.getSpreadSheet().add(spreadSheet);
						spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam);
						data.setReturncode(Constant.SUCCESS_CODE);
						data.setReturndata(excel);
						data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
						data.setMaxRowPixel(returnParam.getMaxRowPixel());
						storeService.set(excelId, excelBook);
					} else {
						data.setReturncode(Constant.CACHE_INVALID_CODE);
						data.setReturndata(Constant.CACHE_INVALID_MSG);
					}
				}else{
					try {
						Thread.sleep(100);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
				}
			}
		}
		if("".equals(data.getReturndata())){
			data.setReturncode(-1);
		}
		this.sendJson(resp, data);
	}
	
	
	
	
	
	
	
	
	
	
	
}
