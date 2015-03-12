/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.InputStream;
import java.io.StringReader;
import java.net.URL;

import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.xml.sax.InputSource;

public class SpreadsheetTestHelper {

	public static WorkbookParser openSpreadsheetResource(String resourceName)
			throws Exception {
		URL resourceURL = SpreadsheetTestHelper.class.getResource(resourceName);
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		return p;
	}

	public static void validateAgainstSchema(String xml) throws Exception {
		URL resource = WorkbookParserXMLTest.class
				.getResource("/schema-v1.xsd");
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
		try {
			reader.read(source);
		} catch (DocumentException e) {
			// System.out.println(xml);
			throw e;
		}

	}

}
