package org.sysmodb;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Namespace;
import org.dom4j.Node;
import org.dom4j.XPath;
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
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testAsDocumentSanity() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		assertNotNull(p.asXMLDocument());
	}

	@Test
	public void testColumnAlphaValues() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		Document doc = p.asXMLDocument();

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		String[] expected = new String[] { "AA", "AB", "BA", "BB", "BC" };
		for (String exp : expected) {
			XPath xpath = DocumentHelper
					.createXPath("//bbb:cell[@column_alpha='" + exp + "']");
			xpath.setNamespaceURIs(namespaceURIs);
			@SuppressWarnings("unchecked")
			List<Node> matches = xpath.selectNodes(doc);
			assertEquals(1, matches.size());
			assertEquals(exp, matches.get(0).getText());
		}
	}

	@SuppressWarnings("unchecked")
	@Test
	public void testNumberOfColumns() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		Document doc = p.asXMLDocument();

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		XPath xpath = DocumentHelper
				.createXPath("//bbb:sheet[@index=\"1\"]//bbb:column");
		xpath.setNamespaceURIs(namespaceURIs);
		List<Node> matches = xpath.selectNodes(doc);
		assertEquals(55, matches.size());
		for (int n = 0; n < matches.size(); n++) {
			assertEquals(String.valueOf(n + 1), matches.get(n)
					.valueOf("@index"));
		}
		xpath = DocumentHelper
				.createXPath("//bbb:sheet[@index=\"1\"]//bbb:columns");
		matches = xpath.selectNodes(doc);
		assertEquals("1", matches.get(0).valueOf("@first_column"));
		assertEquals("55", matches.get(0).valueOf("@last_column"));
	}

	@Test
	public void testFormulaEvaluation() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		Document doc = p.asXMLDocument();

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		XPath xpath = DocumentHelper.createXPath("//bbb:cell[@formula='A1+1']");
		xpath.setNamespaceURIs(namespaceURIs);
		@SuppressWarnings("unchecked")
		List<Node> matches = xpath.selectNodes(doc);
		assertEquals(1, matches.size());
		assertEquals("14.0", matches.get(0).getText());
	}

	@Test
	public void testAsXMLForXLSX() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xlsx");
		assertNotNull(p.asXML());
		SpreadsheetTestHelper.validateAgainstSchema(p.asXML());
	}

	@Test
	public void testDataValidationsXLS() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/simple_annotated_book.xls");
		assertNotNull(p.asXMLDocument());
		String xml = p.asXML();
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testAsDocumentSanityXLSX() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xlsx");
		assertNotNull(p.asXMLDocument());
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

}
