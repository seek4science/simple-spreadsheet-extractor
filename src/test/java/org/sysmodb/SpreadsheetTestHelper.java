/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import static org.junit.Assert.assertNotNull;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.net.URL;

import javax.xml.XMLConstants;
import javax.xml.transform.Source;
import javax.xml.transform.stream.StreamSource;
import javax.xml.validation.Schema;
import javax.xml.validation.SchemaFactory;
import javax.xml.validation.Validator;

import org.xml.sax.SAXException;

public class SpreadsheetTestHelper {

	public static WorkbookParser openSpreadsheetResource(String resourceName) throws Exception {
		URL resourceURL = SpreadsheetTestHelper.class.getResource(resourceName);
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		return p;
	}

	public static void validateAgainstSchema(String xml) throws Exception {
		URL resource = WorkbookParserXMLTest.class.getResource("/schema-v1.xsd");
		try {
            SchemaFactory factory = 
                    SchemaFactory.newInstance(XMLConstants.W3C_XML_SCHEMA_NS_URI);
            Schema schema = factory.newSchema(new File(resource.getFile()));
            Validator validator = schema.newValidator();
            
            Source source = new StreamSource(new StringReader(xml));
            validator.validate(source);
        } catch (IOException | SAXException e) {
        	System.out.println(xml);
            System.out.println("Exception: "+e.getMessage());
            throw e;
        }
		return;

	}

}
