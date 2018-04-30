/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb.xml;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 
 * @author Finn, Stuart Owen
 */
public interface XMLStyleHelper {

	String getBGColour(CellStyle style);

	boolean areFontsEmpty(CellStyle style);

	void writeFontProperties(XMLStreamWriter xmlWriter, CellStyle style) throws XMLStreamException;
}
