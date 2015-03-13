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

import javax.xml.stream.XMLStreamException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.sysmodb.csv.CSVGeneration;
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

	public String asXML() throws IOException, XMLStreamException {
		Writer out = new StringWriter();
		asXML(out);
		return out.toString();
	}

	public void asXML(Writer out) throws IOException, XMLStreamException {

		XMLGeneration generator = new XMLGeneration(poiWorkbook);
		generator.outputToWriter(out);
		out.flush();		
	}

	public String asCSV(int sheetIndex) {
		return asCSV(sheetIndex, false);
	}

	public String asCSV(int sheetIndex, boolean trim) {
		Writer out = new StringWriter();
		asCSV(out, sheetIndex, trim);
		return out.toString();
	}

	public void asCSV(Writer out, int sheetIndex, boolean trim) {
		CSVGeneration generator = new CSVGeneration(poiWorkbook, sheetIndex,
				trim);
		try {
			generator.outputToWriter(out);
			out.flush();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

	public Workbook getWorkbook() {
		return poiWorkbook;
	}

}
