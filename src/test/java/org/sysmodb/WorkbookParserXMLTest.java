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

import org.junit.Test;

public class WorkbookParserXMLTest {

	@Test
	public void testConstruct() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		new WorkbookParser(stream);
	}

	@Test
	public void testValidateXML() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		String xml = p.asXML();
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testValidateAnnotatedXML() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/simple_annotated_book.xls");
		String xml = p.asXML();
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testValidateXLSWithComplexValidations() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/complex_validations.xls");
		String xml = p.asXML();
		System.out.println(xml);
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}
	
	@Test
	public void testAsXMLForXLSX() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xlsx");
		assertNotNull(p.asXML());
		SpreadsheetTestHelper.validateAgainstSchema(p.asXML());
	}
	
	@Test
	public void testJERMTemplatesParsableXLS() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/metabolites_intracellular.xls");
		String xml = p.asXML();
		assertNotNull(xml);
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testJERMTemplatesParsableXLSX() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/metabolites_intracellular.xlsx");
		String xml = p.asXML();
		assertNotNull(xml);
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

        @Test
	public void testValidatedSheet() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/validated_spreadsheet.xls");
		String xml = p.asXML();
                assertNotNull(xml);
                SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

        @Test
        public void testUnknownFormulaSheet() throws Exception {
		//This sheet has STDV.S formula which is known
		//The parser throws error on the cells which the same formula is pulled down
		//Happnes on MS excel 
                WorkbookParser p = SpreadsheetTestHelper
                                .openSpreadsheetResource("/unknown_formula.xlsx");
                String xml = p.asXML();
                assertNotNull(xml);
                SpreadsheetTestHelper.validateAgainstSchema(xml);
        }
}
