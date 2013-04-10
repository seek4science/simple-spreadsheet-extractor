/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import java.io.StringReader;
import java.io.StringWriter;
import java.io.Writer;
import java.lang.reflect.Constructor;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.junit.Test;
import org.xml.sax.InputSource;

public class ControlCharStrippingXMLWriterTest {
	
	/**
	 * Tests first using the normal XMLWriter and checks reading it fails, and then tests using the extended writer using exactly the same xml document,
	 * and checks reading it works.
	 * 
	 * @throws Exception
	 */

	@Test(expected = DocumentException.class)
	public void testXMLWriterFails() throws Exception {
		Document doc = createXMLWithInvalidCharacters();
		String xml = convertWith(doc, XMLWriter.class);		
		readXML(xml);
	}

	@Test
	public void testControlCharStrippingXMLWriter() throws Exception {
		Document doc = createXMLWithInvalidCharacters();
		String xml = convertWith(doc, ControlCharStrippingXMLWriter.class);
		readXML(xml);
	}
	
	
	

	private Document createXMLWithInvalidCharacters() {
		Document doc = DocumentHelper.createDocument();
		Element root = doc.addElement("root");
		root.addAttribute("invalid", "\u0001");

		Element element = root.addElement("inner");
		element.setText("fred\u0003");
		return doc;
	}

	private String convertWith(Document doc, Class<?> writerClass)
			throws Exception {
		StringWriter out = new StringWriter();
		OutputFormat format = OutputFormat.createPrettyPrint();
		format.setEncoding("UTF-8");

		Constructor<?> constructor = writerClass.getConstructor(Writer.class,
				OutputFormat.class);
		XMLWriter writer = (XMLWriter) constructor.newInstance(out, format);

		writer.setEscapeText(true);

		writer.write(doc);
		writer.close();
		String xml = out.toString();
		return xml;
	}

	private void readXML(String xml) throws Exception {
		InputSource source = new InputSource(new StringReader(xml));
		source.setEncoding("UTF-8");
		SAXReader reader = new SAXReader(false);
		reader.read(source);
	}

}
