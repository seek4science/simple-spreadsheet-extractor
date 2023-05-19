/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;

public class CellInfo {

	private final static SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'H:m:sZ");

	public String type;
	public String value;
	public String formula;

	public CellInfo(Cell cell, Workbook workbook) {
		if (cell == null) {
			value = "";
			type = "blank";
			formula = null;
		} else {
			readCellValueAndType(cell.getCellType(), cell);
		}
	}

	private void readCellValueAndType(CellType cellType, Cell cell) {
		switch (cellType) {
		case BLANK:
			value = "";
			type = "blank";
			break;
		case BOOLEAN:
			value = String.valueOf(cell.getBooleanCellValue());
			type = "boolean";
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				type = "datetime";
				Date dateCellValue = cell.getDateCellValue();
				value = dateFormatter.format(dateCellValue);
			} else {
				double numericValue = cell.getNumericCellValue();
				int intValue = (int) numericValue;
				if (intValue == numericValue) {
					value = String.valueOf(intValue);
				} else {
					value = String.valueOf(numericValue);
				}
				type = "numeric";
			}
			break;
		case STRING:
			value = cell.getStringCellValue();
			type = "string";
			break;
		case FORMULA:
			try {
				formula = cell.getCellFormula();
			} catch (FormulaParseException e) {

			}
			CellType resultCellType = cell.getCachedFormulaResultType();
			readCellValueAndType(resultCellType, cell);
			break;
		default:
			value = "";
			type = "none";
			break;
		}
	}
}
