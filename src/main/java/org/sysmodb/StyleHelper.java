/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.ss.usermodel.CellStyle;
import org.dom4j.Element;

 /**
 *
 * @author Finn
 */
public interface StyleHelper {

    String getBGColour(CellStyle style);

    void setFontProperties(CellStyle style, Element element);
    
    boolean areFontsEmpty(CellStyle style);
    
    void writeFontProperties(XMLStreamWriter xmlWriter, CellStyle style) throws XMLStreamException;
}
