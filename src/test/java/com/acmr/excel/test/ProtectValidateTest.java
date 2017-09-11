package com.acmr.excel.test;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Test;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;

import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.copy.Orignal;
import com.acmr.excel.model.copy.Target;
import com.acmr.excel.util.ProtectValidateUtil;
import com.acmr.mq.Model;

public class ProtectValidateTest {
	private ExcelBook excelBook;
	private Model model;

	@Before
	public void before() {
		excelBook = TestUtil.createNewExcel();
		model = new Model();
	}
	
	/**
	 * 测试空单元格受保护条件下的验证
	 */
	@Test
	public void testNullCellProtect1(){
		excelBook.getSheets().get(0).setProtect(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, 2, 2, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试锁定单元格受保护条件下的验证
	 */
	@Test
	public void testLockCellProtect1(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getRows().get(0).set(0, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(0).set(1, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(0).set(2, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(1).set(0, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(1).set(1, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(1).set(2, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(2).set(0, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(2).set(1, new ExcelCell());
		excelBook.getSheets().get(0).getRows().get(2).set(2, new ExcelCell());
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, 2, 2, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试非锁定单元格受保护条件下的验证
	 */
	@Test
	public void testNotLockCellProtect1(){
		excelBook.getSheets().get(0).setProtect(true);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Assert.assertEquals(true, ProtectValidateUtil.validateOpr(0, 0, 2, 2, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试锁定行受保护条件下的验证
	 */
	@Test
	public void testLockRowProtect(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getRows().get(0).getCellstyle().setLocked(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, -1, 0, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试非锁定行受保护条件下的验证1
	 */
	@Test
	public void testNotLockRowProtect(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getRows().get(0).getCellstyle().setLocked(false);
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, -1, 0, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试锁定列受保护条件下的验证
	 */
	@Test
	public void testLockColProtect(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getCols().get(0).getCellstyle().setLocked(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, 0, -1, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试非锁定列受保护条件下的验证1
	 */
	@Test
	public void testNotLockColProtect(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getCols().get(0).getCellstyle().setLocked(false);
		Assert.assertEquals(false, ProtectValidateUtil.validateOpr(0, 0, 0, -1, excelBook.getSheets().get(0)));
	}
	/**
	 * 测试外部粘贴空单元格时的受保护条件下
	 */
	@Test
	public void testNullCellPaste(){
		excelBook.getSheets().get(0).setProtect(true);
		Paste paste = new Paste();
		paste.setOprRow(0);
		paste.setOprCol(0);
		paste.setRowLen(2);;
		paste.setColLen(2);
		model.setReqPath(OperatorConstant.paste);
		model.setObject(paste);
		Assert.assertEquals(false, ProtectValidateUtil.validatePaste(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试外部粘贴锁定单元格时的受保护条件下
	 */
	@Test
	public void testLockPaste(){
		excelBook.getSheets().get(0).setProtect(true);
		Paste paste = new Paste();
		paste.setOprRow(0);
		paste.setOprCol(0);
		paste.setRowLen(2);;
		paste.setColLen(2);
		model.setReqPath(OperatorConstant.paste);
		model.setObject(paste);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(true);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Assert.assertEquals(false, ProtectValidateUtil.validatePaste(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试外部粘贴非锁定单元格时的受保护条件下
	 */
	@Test
	public void testNotLockPaste(){
		excelBook.getSheets().get(0).setProtect(true);
		Paste paste = new Paste();
		paste.setOprRow(0);
		paste.setOprCol(0);
		paste.setRowLen(2);;
		paste.setColLen(2);
		model.setReqPath(OperatorConstant.paste);
		model.setObject(paste);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Assert.assertEquals(true, ProtectValidateUtil.validatePaste(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试外部粘贴请求行数大于实际行数受保护条件下
	 */
	@Test
	public void testMoreRowPaste(){
		excelBook.getSheets().get(0).setProtect(true);
		Paste paste = new Paste();
		paste.setOprRow(102);
		paste.setOprCol(0);
		paste.setRowLen(2);;
		paste.setColLen(2);
		model.setReqPath(OperatorConstant.paste);
		model.setObject(paste);
		Assert.assertEquals(false, ProtectValidateUtil.validatePaste(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试外部粘贴请求行数大于实际行数受保护条件下
	 */
	@Test
	public void testMoreColPaste(){
		excelBook.getSheets().get(0).setProtect(true);
		Paste paste = new Paste();
		paste.setOprRow(0);
		paste.setOprCol(102);
		paste.setRowLen(2);;
		paste.setColLen(2);
		model.setReqPath(OperatorConstant.paste);
		model.setObject(paste);
		Assert.assertEquals(false, ProtectValidateUtil.validatePaste(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试锁定下的复制粘贴
	 */
	@Test
	public void testLockCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(0);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(true);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试非锁定下的复制粘贴
	 */
	@Test
	public void testNotLockCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(0);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Assert.assertEquals(true, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试目标行为非锁定下的复制粘贴
	 */
	@Test
	public void testTargetNotLockCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(4);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		ExcelCell newCell = new ExcelCell();
		newCell.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(4).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(4).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(4).set(2, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(2, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(2, newCell);
		Assert.assertEquals(true, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试目标行为锁定下的复制粘贴
	 */
	@Test
	public void testTargetLockCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(4);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		ExcelCell newCell = new ExcelCell();
		newCell.getCellstyle().setLocked(true);
		excelBook.getSheets().get(0).getRows().get(4).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(4).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(4).set(2, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(5).set(2, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(0, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(1, newCell);
		excelBook.getSheets().get(0).getRows().get(6).set(2, newCell);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试非锁定下的复制粘贴,原始行大于实际行
	 */
	@Test
	public void testMoreRowCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(104);
		orignal.setStartCol(0);
		orignal.setStartRow(102);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(0);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试非锁定下的复制粘贴,原始列大于实际列
	 */
	@Test
	public void testMoreColCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(104);
		orignal.setEndRow(2);
		orignal.setStartCol(102);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(0);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试非锁定下的复制粘贴,目标行大于实际行
	 */
	@Test
	public void testMoreTargetRowCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Target target = new Target();
		target.setOprCol(0);
		target.setOprRow(102);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试非锁定下的复制粘贴,原始列大于实际列
	 */
	@Test
	public void testMoreTargetColCopyOrCut(){
		excelBook.getSheets().get(0).setProtect(true);
		Copy copy = new Copy();
		Orignal orignal = new Orignal();
		orignal.setEndCol(2);
		orignal.setEndRow(2);
		orignal.setStartCol(0);
		orignal.setStartRow(0);
		copy.setOrignal(orignal);
		ExcelCell cell00 = new ExcelCell();
		cell00.getCellstyle().setLocked(false);
		excelBook.getSheets().get(0).getRows().get(0).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(0).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(1).set(2, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(0, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(1, cell00);
		excelBook.getSheets().get(0).getRows().get(2).set(2, cell00);
		Target target = new Target();
		target.setOprCol(100);
		target.setOprRow(0);
		copy.setTarget(target);
		model.setReqPath(OperatorConstant.cut);
		model.setObject(copy);
		Assert.assertEquals(false, ProtectValidateUtil.validateCopyOrCut(model, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试行锁定的受保护条件下
	 */
	@Test
	public void testLockRow(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getRows().get(0).getCellstyle().setLocked(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateRow(0, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试行非锁定的受保护条件下
	 */
	@Test
	public void testNotLockRow(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getRows().get(0).getCellstyle().setLocked(false);
		Assert.assertEquals(true, ProtectValidateUtil.validateRow(0, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试行大于实际行数的受保护条件下
	 */
	@Test
	public void testLockMoreRow(){
		excelBook.getSheets().get(0).setProtect(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateRow(102, excelBook.getSheets().get(0).getRows()));
	}
	/**
	 * 测试列锁定的受保护条件下
	 */
	@Test
	public void testLockCol(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getCols().get(0).getCellstyle().setLocked(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateCol(0, excelBook.getSheets().get(0).getCols()));
	}
	/**
	 * 测试列非锁定的受保护条件下
	 */
	@Test
	public void testNotLockCol(){
		excelBook.getSheets().get(0).setProtect(true);
		excelBook.getSheets().get(0).getCols().get(0).getCellstyle().setLocked(false);
		Assert.assertEquals(true, ProtectValidateUtil.validateCol(0, excelBook.getSheets().get(0).getCols()));
	}
	/**
	 * 测试列大于实际列数的受保护条件下
	 */
	@Test
	public void testLockMoreCol(){
		excelBook.getSheets().get(0).setProtect(true);
		Assert.assertEquals(false, ProtectValidateUtil.validateCol(102, excelBook.getSheets().get(0).getCols()));
	}
	
	/**
	 * 保护，文本内容
	 */
	@Test
	public void testValidateStatus(){
		excelBook.getSheets().get(0).setProtect(true);
		Assert.assertEquals(true, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.textData));
		Assert.assertEquals(true, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.paste));
		Assert.assertEquals(true, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.copy));
		Assert.assertEquals(true, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.cut));
		Assert.assertEquals(false, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.addRowLine));
		excelBook.getSheets().get(0).setProtect(false);
		Assert.assertEquals(false, ProtectValidateUtil.validateStatus(excelBook.getSheets().get(0), OperatorConstant.copy));
	}
	
	
	
}
