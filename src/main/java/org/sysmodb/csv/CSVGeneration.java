package org.sysmodb.csv;

import java.io.IOException;
import java.io.Writer;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.sysmodb.CellInfo;

public class CSVGeneration {
	private final int sheetIndex;
	private final Workbook poiWorkbook;
	private final boolean trim;

	public CSVGeneration(Workbook poiWorkbook, int sheetIndex, boolean trim) {
		this.poiWorkbook = poiWorkbook;
		this.sheetIndex = sheetIndex;
		this.trim = trim;
	}

	public void outputToWriter(Writer outputWriter) throws IOException {
		// FIXME: this is a fairly naive implementation, just to get something
		// working quickly.
		// For a better implementation should do something event driven like:
		// https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java

		Sheet sheet = poiWorkbook.getSheetAt(sheetIndex - 1);
		int lastRow = sheet.getLastRowNum();
		int lastCol = -1;
		int firstRow = 0;
		int firstCol = 0;
		List<String> stringRowTypes = Arrays.asList(new String[] { "string", "datetime" });

		if (trim) {
			firstRow = sheet.getFirstRowNum();
			firstCol = 257;
		}

		for (int i = sheet.getFirstRowNum(); i <= lastRow; i++) {
			Row row = sheet.getRow(i);
			if (row != null && row.getLastCellNum() > lastCol) {
				lastCol = row.getLastCellNum();
			}
			if (trim && row != null && row.getFirstCellNum() < firstCol) {
				firstCol = row.getFirstCellNum();
			}
		}

		String blankRow = "";
		for (int i = 0; i < lastCol - 1; i++) {
			blankRow += ",";
		}

		for (int y = firstRow; y <= lastRow; y++) {
			Row row = sheet.getRow(y);
			String csvRow = "";
			if (row != null) {
				for (int x = firstCol; x < lastCol; x++) {
					Cell cell = row.getCell(x);
					CellInfo info = new CellInfo(cell, poiWorkbook);
					String value = info.value;
					if (info.type.equalsIgnoreCase("boolean"))
						value = value.toUpperCase();
					if (stringRowTypes.contains(info.type))
						value = "\"" + value + "\"";
					csvRow += value;
					if (x != lastCol - 1)
						csvRow += ",";
				}
				outputWriter.write(csvRow);
			} else {
				outputWriter.write(blankRow);
			}

			if (y != lastRow)
				outputWriter.write("\n");
		}
	}
}
