/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.StringWriter;
import java.io.Writer;
import java.net.URL;

import org.junit.Ignore;
import org.junit.Test;

public class WorkbookParserCSVTest {

	@Test
	@Ignore("Problem with the test rather than what it is testing. Seems to be a charset problem")
	public void testAsCSV() throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class
				.getResource("/test-spreadsheet-for-csv.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String csv = p.asCSV(1);
		assertNotNull(csv);

		String expectedResult = expectedResult("/test-spreadsheet-for-csv.csv");
		assertEquals(expectedResult, csv);
	}

	@Test
	public void testAsCSVAnotherSheet() throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class
				.getResource("/test-spreadsheet-for-csv.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String csv = p.asCSV(2);
		assertNotNull(csv);

		assertEquals(",,\"a\",1,TRUE,,FALSE", csv);
	}

	@Test
	public void testCSVWithBlankRow() throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class
				.getResource("/test-spreadsheet-for-csv.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String csv = p.asCSV(3);
		assertNotNull(csv);

		String expectedResult = expectedResult("/test-spreadsheet-for-csv-blank-row.csv");
		assertEquals(expectedResult, csv);
	}

	@Test
	public void testAsCSVTrimmed() throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class
				.getResource("/test-spreadsheet-for-csv.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String csv = p.asCSV(1, true);
		assertNotNull(csv);

		String expectedResult = expectedResult("/test-spreadsheet-for-csv-trimmed.csv");
		assertEquals(expectedResult, csv);
	}

	private String expectedResult(String resourceName) throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class.getResource(resourceName);

		Writer writer = new StringWriter();

		char[] buffer = new char[1024];

		Reader reader = new BufferedReader(new InputStreamReader(
				resourceURL.openStream(), "UTF-8"));
		int n;
		while ((n = reader.read(buffer)) != -1) {
			writer.write(buffer, 0, n);
		}
		return writer.toString();
	}
}
