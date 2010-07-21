package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.net.URL;

import org.junit.Ignore;
import org.junit.Test;

public class WorkbookParserTest {

	@Test
	public void testConstruct() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		new WorkbookParser(stream);
	}
	
	@Test
	public void testAsXMLSanity() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXML());
		System.out.println(p.asXML());
	}
	
	@Test
	public void testAsDocumentSanity() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXMLDocument());
	}

	@Test
	@Ignore
	public void testConstruct2() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		new WorkbookParser(stream);
	}
}
