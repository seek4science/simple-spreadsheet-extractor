package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.InputStream;
import java.io.StringBufferInputStream;
import java.net.URL;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.dom4j.Document;
import org.dom4j.io.SAXReader;
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
	}

	@Test
	public void testValidateXML() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String xml = p.asXML();
		validateAgainstSchema(xml);
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
	public void testAsXMLSanityXLSX() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXML());
	}

	@Test
	public void testAsDocumentSanityXLSX() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		assertNotNull(p.asXMLDocument());
	}

	@Test
	public void testConstructXLSX() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		new WorkbookParser(stream);
	}

	@Test
	public void testvalidateXLSXXML() throws Exception {
		URL resourceURL = WorkbookParserTest.class
				.getResource("/test-spreadsheet.xlsx");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String xml = p.asXML();
		validateAgainstSchema(xml);
	}

	private void validateAgainstSchema(String xml) throws Exception {
		System.out.println(xml);
		URL schemaURL = WorkbookParserTest.class.getResource("/schema-v1.xsd");
		assertNotNull(schemaURL);
		SAXParserFactory factory = SAXParserFactory.newInstance();
		factory.setValidating(true);
		factory.setNamespaceAware(true);

		SAXParser parser = factory.newSAXParser();
		parser.setProperty(
				"http://java.sun.com/xml/jaxp/properties/schemaLanguage",
				"http://www.w3.org/2001/XMLSchema");
		parser.setProperty(
				"http://java.sun.com/xml/jaxp/properties/schemaSource",
				schemaURL.toString());

		SAXReader reader = new SAXReader(parser.getXMLReader());
		Document doc = reader.read(new StringBufferInputStream(xml));

	}
}
