package com.acmr.excel.controller;

import java.io.IOException;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.spy.memcached.MemcachedClient;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.service.PasteService;

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
	private MemcachedClient memcachedClient;
	@Resource
	private PasteService pasteService; 

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
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(req.getHeader("excelId"));
		boolean isAblePasteResult = pasteService.isAblePaste(paste, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, paste, OperatorConstant.paste);
		}else{
			this.assemblePasteData(req, resp);
		}
		
	}
	/**
	 * 内部复制粘贴
	 * @throws IOException
	 */
	@RequestMapping("/copy")
	public void copy(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(req.getHeader("excelId"));
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.copy);
		}else{
			this.assemblePasteData(req, resp);
		}
	}
	/**
	 * 剪切粘贴
	 * @throws IOException
	 */
	@RequestMapping("/cut")
	public void cut(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)memcachedClient.get(copy.getExcelId());
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.cut);
		}else{
			this.assemblePasteData(req, resp);
		}
	}
}
