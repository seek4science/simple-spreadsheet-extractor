/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.net.URL;

import javax.xml.stream.XMLStreamException;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExtractorMain {

	private static OptionParser options;
	private static final Logger logger = Logger.getLogger(ExtractorMain.class);

	public ExtractorMain(String[] args) {
		processOptions(args);
		try {
			WorkbookParser workbook = new WorkbookParser(getInputStream());
			if (options.getOutputFormat().equals("xml")) {
				workbook.asXML(new OutputStreamWriter(System.out));
			}
			if (options.getOutputFormat().equals("csv")) {
				workbook.asCSV(new OutputStreamWriter(System.out), options.getSheet(), options.getTrim());
			}

		} catch (IOException e) {
			logger.error("IO Error reading data: ", e);
			System.exit(-1);
		} catch (InvalidFormatException e) {
			logger.error("Invaild format reading data", e);
			System.exit(-1);
		} catch (XMLStreamException e) {
			logger.error("Error writing to xml stream", e);
			System.exit(-1);
		}
		System.exit(0);
	}

	public static void main(String[] args) {
		new ExtractorMain(args);
	}

	private InputStream getInputStream() throws IOException {
		if (options.getFilename() != null) {
			URL url = new URL("file://" + options.getFilename());
			return url.openStream();
		} else { // get from stdin
			return System.in;
		}
	}

	private void processOptions(String[] args) {
		try {
			options = new OptionParser(args);
		} catch (InvalidOptionException e) {
			System.err.println("Invalid option: " + e.getMessage());
			System.exit(-1);
		}
	}

}
