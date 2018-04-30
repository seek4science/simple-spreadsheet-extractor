/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;

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
	private static final Logger logger = Logger.getLogger("SpreadsheetSuiteTest");
	private final boolean FAIL_QUIETLY = false;

	/**
	 * Tests each spreadsheet resource defined by
	 * {@link #getSpreadsheetResourceNames()}
	 */
	@Test
	public void testAll() throws Exception {

		for (String name : getSpreadsheetResourceNames()) {
			System.out.println(name);
			try {
				URL resourceURL = WorkbookParserXMLTest.class.getResource("/" + ROOT + "/" + name);
				assertNotNull(resourceURL);
				InputStream stream = resourceURL.openStream();
				WorkbookParser p = new WorkbookParser(stream);
				String xml = p.asXML();
				SpreadsheetTestHelper.validateAgainstSchema(xml);
			} catch (Exception e) {
				logger.severe("Error parsing " + name + " - " + e.getMessage());
				if (!FAIL_QUIETLY) {
					throw e;
				}
			}
		}
	}

	private List<String> getSpreadsheetResourceNames() {
		List<String> names = new ArrayList<String>();
		names.add("problematic_spreadsheet.xls");
		names.add("problematic_spreadsheet2.xls");
		names.add("problematic_spreadsheet3.xls");
		names.add("pride_template_empty.xls");
		names.add("xml-unfriendly-chars.xlsx");
		names.add("chars-that-need-escaping.xlsx");
		return names;
	}

}
