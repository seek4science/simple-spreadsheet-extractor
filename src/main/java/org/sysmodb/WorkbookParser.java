/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;
import java.io.Writer;
import java.util.Arrays;
import java.util.List;

import javax.xml.stream.XMLStreamException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.sysmodb.xml.XMLGeneration;

/**
 * 
 * @author Stuart Owen, Finn Bacall
 * 
 */
public class WorkbookParser {	

	private Workbook poiWorkbook = null;	
		
	
	public WorkbookParser(InputStream stream) throws IOException,
			InvalidFormatException {

		poiWorkbook = WorkbookFactory.create(stream);
		
	}
	
	public String asXML() {
		Writer out = new StringWriter();
		asXML(out);
		return out.toString();
	}
	
	public void asXML(Writer out) {
		
		XMLGeneration generator = new XMLGeneration(poiWorkbook);
		try {
			generator.outputToWriter(out);
		} catch (IOException e) {			
			e.printStackTrace();			
		} catch (XMLStreamException e) {
			e.printStackTrace();			
		} 
		
	}

	public String asCSV(int sheetIndex) {
		return asCSV(sheetIndex, false);
	}

	public String asCSV(int sheetIndex, boolean trim) {
		// FIXME: this is a fairly naive implementation, just to get something
		// working quickly.
		// For a better implementation should do something event driven like:
		// https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java
		String result = "";
		Sheet sheet = poiWorkbook.getSheetAt(sheetIndex - 1);
		int lastRow = sheet.getLastRowNum();
		int lastCol = -1;
		int firstRow = 0;
		int firstCol = 0;
		List<String> stringRowTypes = Arrays.asList(new String[] { "string",
				"datetime" });

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
					CellInfo info = new CellInfo(cell,poiWorkbook);
					String value = info.value;
					if (info.type.equalsIgnoreCase("boolean"))
						value = value.toUpperCase();
					if (stringRowTypes.contains(info.type))
						value = "\"" + value + "\"";
					csvRow += value;
					if (x != lastCol - 1)
						csvRow += ",";
				}
				result += csvRow;
			} else {
				result += blankRow;
			}

			if (y != lastRow)
				result += "\n";
		}

		return result;
	}
	
	public Workbook getWorkbook() {
		return poiWorkbook;
	}

		
}
