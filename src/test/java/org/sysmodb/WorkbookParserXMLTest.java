package org.sysmodb;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.InputStream;
import java.io.StringReader;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Namespace;
import org.dom4j.Node;
import org.dom4j.XPath;
import org.dom4j.io.SAXReader;
import org.junit.Test;
import org.xml.sax.InputSource;

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
	public void testAsXMLSanity() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXML());
	}

	@Test
	public void testValidateXML() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String xml = p.asXML();
		validateAgainstSchema(xml);
	}
	
	@Test
	public void testValidateAnnotatedXML() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/simple_annotated_book.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String xml = p.asXML();
		validateAgainstSchema(xml);
	}

	@Test
	public void testAsDocumentSanity() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXMLDocument());
	}

	@Test
	public void testColumnAlphaValues() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		Document doc = p.asXMLDocument();
		System.out.println(p.asXML());

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		String[] expected = new String[] { "AA", "AB", "BA", "BB", "BC" };
		for (String exp : expected) {
			XPath xpath = DocumentHelper
					.createXPath("//bbb:cell[@column_alpha='" + exp + "']");
			xpath.setNamespaceURIs(namespaceURIs);
			List<Node> matches = xpath.selectNodes(doc);
			assertEquals(1, matches.size());
			assertEquals(exp, matches.get(0).getText());
		}
	}
	
  @Test
  public void testNumberOfColumns() throws Exception {
    URL resourceURL = WorkbookParserXMLTest.class
        .getResource("/test-spreadsheet.xls");
    assertNotNull(resourceURL);
    InputStream stream = resourceURL.openStream();
    WorkbookParser p = new WorkbookParser(stream);
    Document doc = p.asXMLDocument();
    System.out.println(p.asXML());

    Namespace defNamespace = doc.getRootElement().getNamespace();
    doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
    Map<String, String> namespaceURIs = new HashMap<String, String>();
    namespaceURIs.put("bbb", defNamespace.getURI());
    XPath xpath = DocumentHelper
        .createXPath("//bbb:sheet[@index=\"1\"]//bbb:column");
    xpath.setNamespaceURIs(namespaceURIs);
    List<Node> matches = xpath.selectNodes(doc);
    assertEquals(55, matches.size());
    for (int n = 0; n < matches.size(); n++)
    {
      assertEquals(String.valueOf(n+1), matches.get(n).valueOf("@index"));
    }
    xpath = DocumentHelper.createXPath("//bbb:sheet[@index=\"1\"]//bbb:columns");
    matches = xpath.selectNodes(doc);
    assertEquals("1", matches.get(0).valueOf("@first_column"));
    assertEquals("55", matches.get(0).valueOf("@last_column"));
  }

	@Test
	public void testFormulaEvaluation() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		Document doc = p.asXMLDocument();

		Namespace defNamespace = doc.getRootElement().getNamespace();
		doc.getRootElement().addNamespace("bbb", defNamespace.getURI());
		Map<String, String> namespaceURIs = new HashMap<String, String>();
		namespaceURIs.put("bbb", defNamespace.getURI());
		XPath xpath = DocumentHelper.createXPath("//bbb:cell[@formula='A1+1']");
		xpath.setNamespaceURIs(namespaceURIs);
		List<Node> matches = xpath.selectNodes(doc);
		assertEquals(1, matches.size());
		assertEquals("14.0", matches.get(0).getText());
	}

	@Test
	public void testAsXMLSanityXLSX() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXML());
	}

	@Test
	public void testDataValidationsXLS() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
		.getResource("/simple_annotated_book.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		System.out.println(p.asXML());
	}
	
	@Test
	public void testAsDocumentSanityXLSX() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXMLDocument());
	}
	
  @Test
  public void testJERMTemplatesParsableXLS() throws Exception {
    URL resourceURL = WorkbookParserXMLTest.class
        .getResource("/metabolites_intracellular.xls");
    assertNotNull(resourceURL);
    InputStream stream = resourceURL.openStream();
    WorkbookParser p = new WorkbookParser(stream);
    assertNotNull(p.asXML());
  }
  
  @Test
  public void testJERMTemplatesParsableXLSX() throws Exception {
    URL resourceURL = WorkbookParserXMLTest.class
        .getResource("/metabolites_intracellular.xls");
    assertNotNull(resourceURL);
    InputStream stream = resourceURL.openStream();
    WorkbookParser p = new WorkbookParser(stream);
    assertNotNull(p.asXML());
  }

	@Test
	public void testConstructXLSX() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		new WorkbookParser(stream);
	}

	@Test
	public void testvalidateXLSXXML() throws Exception {
		URL resourceURL = WorkbookParserXMLTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String xml = p.asXML();
		validateAgainstSchema(xml);
	}

	private void validateAgainstSchema(String xml) throws Exception {
		URL resource = WorkbookParserXMLTest.class.getResource("/schema-v1.xsd");
		SAXReader reader = new SAXReader(true);
		reader.setFeature("http://apache.org/xml/features/validation/schema",
				true);
		reader.setProperty(
				"http://java.sun.com/xml/jaxp/properties/schemaLanguage",
				"http://www.w3.org/2001/XMLSchema");
		reader.setProperty(
				"http://java.sun.com/xml/jaxp/properties/schemaSource",
				new File(resource.getFile()));
		InputSource source = new InputSource(new StringReader(xml));
		source.setEncoding("UTF-8");
		reader.read(source);
	}
}
