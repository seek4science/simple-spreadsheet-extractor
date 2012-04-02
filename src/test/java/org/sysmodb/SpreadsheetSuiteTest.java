package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

/**
 * A test suite made up of a collection of spreadsheets that have failed to be
 * processed in the past, for various reasons, and aren't tied to a specific
 * test case.
 * 
 * @author Stuart Owen
 */
public class SpreadsheetSuiteTest {
	private final String ROOT = "spreadsheet_suite";

	/**
	 * Tests each spreadsheet resource defined by
	 * {@link #getSpreadsheetResourceNames()}
	 */
	@Test
	public void testAll() throws Exception {

		for (String name : getSpreadsheetResourceNames()) {
			URL resourceURL = WorkbookParserXMLTest.class.getResource("/"
					+ ROOT + "/" + name);
			assertNotNull(resourceURL);
			InputStream stream = resourceURL.openStream();
			WorkbookParser p = new WorkbookParser(stream);
			String xml = p.asXML();
			SpreadsheetTestHelper.validateAgainstSchema(xml);

		}
	}

	private List<String> getSpreadsheetResourceNames() {
		return new ArrayList<String>();
	}

}
