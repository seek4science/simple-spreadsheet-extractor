package org.sysmodb.xml;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.IOException;
import java.io.StringWriter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.stream.XMLStreamException;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Namespace;
import org.dom4j.Node;
import org.dom4j.XPath;
import org.junit.Test;
import org.sysmodb.SpreadsheetTestHelper;
import org.sysmodb.WorkbookParser;

public class XMLGenerationTest {

	@Test
	public void testAsDocumentSanity() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		assertNotNull(convertToXMLDocument(new XMLGeneration(p.getWorkbook())));
	}

	@Test
	public void testDataValidationsXLS() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/simple_annotated_book.xls");
		assertNotNull(convertToXMLDocument(new XMLGeneration(p.getWorkbook())));
		String xml = p.asXML();
		SpreadsheetTestHelper.validateAgainstSchema(xml);
	}

	@Test
	public void testAsDocumentSanityXLSX() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xlsx");
		assertNotNull(convertToXMLDocument(new XMLGeneration(p.getWorkbook())));
	}

	@Test
	public void testNumberFormatting() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/numbers_and_strings.xls");
		Document doc = convertToXMLDocument(new XMLGeneration(p.getWorkbook()));
		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		String[] expected = new String[] { "49", "49", "49", "49.95", "49.95",
				"49.95" };
		XPath xpath = DocumentHelper
				.createXPath("//bbb:cell[@column_alpha='B']");
		@SuppressWarnings("unchecked")
		List<Node> nodes = xpath.selectNodes(doc);
		int i = 0;
		for (Node node : nodes) {
			String val = node.getStringValue();
			assertEquals(expected[i++], val);
		}
	}

	@Test
	public void testColumnAlphaValues() throws Exception {
		WorkbookParser p = SpreadsheetTestHelper
				.openSpreadsheetResource("/test-spreadsheet.xls");
		Document doc = convertToXMLDocument(new XMLGeneration(p.getWorkbook()));

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
		Document doc = convertToXMLDocument(new XMLGeneration(p.getWorkbook()));

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
		Document doc = convertToXMLDocument(new XMLGeneration(p.getWorkbook()));

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		XPath xpath = DocumentHelper.createXPath("//bbb:cell[@formula='A1+1']");
		xpath.setNamespaceURIs(namespaceURIs);
		@SuppressWarnings("unchecked")
		List<Node> matches = xpath.selectNodes(doc);
		assertEquals(1, matches.size());
		assertEquals("14", matches.get(0).getText());
	}

	private Document convertToXMLDocument(XMLGeneration generator)
			throws IOException, DocumentException, XMLStreamException {
		StringWriter out = new StringWriter();
		generator.outputToWriter(out);
		Document document = DocumentHelper.parseText(out.toString());
		// printOutXML(document);
		return document;
	}

	// private void printOutXML(Document document) {
	// StringWriter out = new StringWriter();
	// OutputFormat format = OutputFormat.createPrettyPrint();
	// format.setEncoding("UTF-8");
	// XMLWriter writer = new XMLWriter(out, format);
	//
	// writer.setEscapeText(true);
	//
	// try {
	// writer.write(document);
	// writer.close();
	// System.out.println(out.toString());
	// } catch (IOException e) {
	// // should never get here, since we are using a StringWriter rather
	// // than IO based Writer
	// e.printStackTrace();
	// }
	// }
}
