package org.sysmodb;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class CellInfo {
	
	private final static SimpleDateFormat dateFormatter = new SimpleDateFormat(
			"yyyy-MM-dd'T'H:m:sZ");
	
	public String type;
	public String value;
	public String formula;
	
	public CellInfo(Cell cell, Workbook workbook) {
		if (cell == null) {
			value = "";
			type = "blank";
			formula = null;
		} else {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				value = "";
				type = "blank";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				value = String.valueOf(cell.getBooleanCellValue());
				type = "boolean";
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					type = "datetime";
					Date dateCellValue = cell.getDateCellValue();
					value = dateFormatter.format(dateCellValue);
				} else {
					value = String.valueOf(cell.getNumericCellValue());
					type = "numeric";
				}
				break;
			case Cell.CELL_TYPE_STRING:
				value = cell.getStringCellValue();
				type = "string";
				break;
			case Cell.CELL_TYPE_FORMULA:

				FormulaEvaluator evaluator = workbook.getCreationHelper()
						.createFormulaEvaluator();
				CellValue cellValue;
				formula = cell.getCellFormula();
				try {
					cellValue = evaluator.evaluate(cell);
					value = cellValue.formatAsString();					
					switch (cellValue.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						type = "boolean";
						break;
					case Cell.CELL_TYPE_STRING:
						type = "string";
						break;
					case Cell.CELL_TYPE_NUMERIC:
						type = "numeric";
						if (DateUtil.isCellDateFormatted(cell)) {
							type = "datetime";
							value = dateFormatter.format(DateUtil
									.getJavaDate(cellValue.getNumberValue()));
						}
						break;
					}
				}
				catch(NotImplementedException e) {
					type="error";
					value=e.getMessage();
				}
				
				break;
			}
		}
	}
}
