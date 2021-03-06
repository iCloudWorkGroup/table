package com.acmr.excel.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;












import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.returnData;

import com.acmr.cache.MemoryUtil;
import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.AreaSet;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.Protect;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.ReturnData;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.datavalidate.Data;
import com.acmr.excel.model.datavalidate.Rule;
import com.acmr.excel.model.position.OpenExcel;
import com.acmr.excel.service.ExcelService;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.SheetService;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.util.AnsycDataReturn;
import com.acmr.excel.util.DataValidateUtil;
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
	@Resource
	private SheetService sheetService;

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
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Paste paste = getJsonDataParameter(req, Paste.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(excelId);
		boolean isAblePasteResult = pasteService.isAblePaste(paste, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, paste, OperatorConstant.paste);
			ansycDataReturn.setIsLegal(true);
		}else{
			ansycDataReturn.setIsLegal(false);
		}
		this.sendJson(resp, ansycDataReturn);
	}
	/**
	 * 内部复制粘贴
	 * @throws IOException
	 */
	@RequestMapping("/copy")
	public void copy(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(excelId);
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.copy);
			ansycDataReturn.setIsLegal(true);
		}else{
			ansycDataReturn.setIsLegal(false);
		}
		this.sendJson(resp, ansycDataReturn);
	}
	/**
	 * 剪切粘贴
	 * @throws IOException
	 */
	@RequestMapping("/cut")
	public void cut(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)storeService.get(excelId);
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.cut);
			ansycDataReturn.setIsLegal(true);
		}else{
			ansycDataReturn.setIsLegal(false);
		}
		this.sendJson(resp, ansycDataReturn);
	}
	
	
	/**
	 * 通过像素动态加载excel
	 */
	@RequestMapping("/area")
	public void openexcel(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		String excelId = req.getHeader("excelId");
		String curStep = req.getHeader("step");
		JsonReturn data = new JsonReturn("");
		if(StringUtil.isEmpty(excelId) || StringUtil.isEmpty(curStep)){
			resp.setStatus(400);
			return;
		}
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		int memStep = (int)storeService.get(excelId+"_ope");
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		ExcelBook excelBook = (ExcelBook) storeService.get(excelId);
		
		if (cStep == memStep) {
			if (excelBook != null) {
				ExcelSheet excelSheet = excelBook.getSheets().get(0);
				ReturnParam returnParam = new ReturnParam();
				CompleteExcel excel = new CompleteExcel();
				SpreadSheet spreadSheet = new SpreadSheet();
				excel.getSpreadSheet().add(spreadSheet);
				spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam,excelId);
				data.setReturncode(Constant.SUCCESS_CODE);
				data.setReturndata(excel);
				data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
				data.setMaxRowPixel(returnParam.getMaxRowPixel());
				//storeService.set(excelId,excelBook);
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
						spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam,excelId);
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
	
	/**
	 * sheet保护
	 * @param req
	 * @param resp
	 * @throws IOException
	 */
	@RequestMapping("/protect")
	public void protect(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		String excelId = req.getHeader("excelId");
		Protect protect = getJsonDataParameter(req, Protect.class);
		ExcelBook excelBook = (ExcelBook) storeService.get(excelId);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(!protect.isProtect()){
			if (excelBook.getSheets().get(0).getPassword() == null || 
					excelBook.getSheets().get(0).getPassword().equals(protect.getPassword())) {
				ansycDataReturn.setIsLegal(true);
				this.assembleData(req, resp,protect,OperatorConstant.PROTECT);
			}else{
				ansycDataReturn.setIsLegal(false);
			}
			this.sendJson(resp, ansycDataReturn);
		}else{
			this.assembleData(req, resp,protect,OperatorConstant.PROTECT);
		}
	}
	/**
	 * sheet保护
	 * @param req
	 * @param resp
	 * @throws IOException
	 */
	@RequestMapping("/validate-set")
	public void lock(HttpServletRequest req, HttpServletResponse resp) throws Exception {
    	AreaSet cell = getJsonDataParameter(req, AreaSet.class);
		this.assembleData(req, resp, cell, OperatorConstant.DATAVALIDATE);
	}
	@RequestMapping("/validate-full")
	public void validateFull(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		String excelId = req.getHeader("excelId");
    	Frozen frozen = getJsonDataParameter(req, Frozen.class);
    	ExcelBook excelBook = (ExcelBook)storeService.get(excelId);
    	ExcelSheet excelSheet = excelBook.getSheets().get(0);
		Data data = MemoryUtil.getDataValidateMap().get(excelId);
		int ruleIndex = sheetService.getRule(data, frozen, excelSheet);
		Rule rule = data.getRuleList().get(ruleIndex);
		ReturnData ReturnData = new ReturnData();
		List<String> seqList = sheetService.findSeq(rule.getFormula1(), excelBook.getSheets().get(0));
		Rule newRule = new Rule();
		final String formal = rule.getFormula1();
		newRule.setFormula1(DataValidateUtil.alias2Display(formal, excelSheet));
		newRule.setValidationType(rule.getValidationType());
		ReturnData.setRule(newRule);
		ReturnData.setIndex(ruleIndex);
		ReturnData.setExpResult(seqList);
		this.sendJson(resp, ReturnData);
	}
}
