/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb.xml;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 *
 * @author Finn
 */
public abstract class XMLStyleGenerator {
  private static final Map<Short, String> BORDERS = createBorderMap();

  private static Map<Short, String> createBorderMap() {
    Map<Short, String> result = new HashMap<Short, String>();
    result.put(CellStyle.BORDER_DASHED, "dashed 1pt");
    result.put(CellStyle.BORDER_DASH_DOT, "dashed 1pt");
    result.put(CellStyle.BORDER_DASH_DOT_DOT, "dashed 1pt");
    result.put(CellStyle.BORDER_DOTTED, "dotted 1pt");
    result.put(CellStyle.BORDER_DOUBLE, "double 3pt");
    result.put(CellStyle.BORDER_HAIR, "solid 1pt");
    result.put(CellStyle.BORDER_MEDIUM, "2pt solid");
    result.put(CellStyle.BORDER_MEDIUM_DASHED, "2pt dashed");
    result.put(CellStyle.BORDER_MEDIUM_DASH_DOT, "2pt dashed");
    result.put(CellStyle.BORDER_MEDIUM_DASH_DOT_DOT, "2pt dashed");
    result.put(CellStyle.BORDER_NONE, "none");
    result.put(CellStyle.BORDER_SLANTED_DASH_DOT, "dashed 2pt");
    result.put(CellStyle.BORDER_THICK, "solid 3pt");
    result.put(CellStyle.BORDER_THIN, "dashed 1pt");

    return Collections.unmodifiableMap(result);
  }

  public static boolean isStyleEmpty(CellStyle style,XMLStyleHelper helper) {
	  
	  if(BORDERS.get(style.getBorderTop()) != "none")
	      return false;
	    if(BORDERS.get(style.getBorderBottom()) != "none")
	    	return false;
	    if(BORDERS.get(style.getBorderLeft()) != "none")
	    	return false;
	    if(BORDERS.get(style.getBorderRight()) != "none")
	    	return false;
	  
	    //Background/fill colour	    
	    if((helper.getBGColour(style)) != null) 
	    	return false;
	    
	    return helper.areFontsEmpty(style);	    
  }
  
  public static void writeStyle(XMLStreamWriter xmlWriter, CellStyle style, XMLStyleHelper helper) throws XMLStreamException {
	  String border = "none";
	  xmlWriter.writeStartElement("style");
	  xmlWriter.writeAttribute("id", "style"+style.getIndex());
	    if((border = BORDERS.get(style.getBorderTop())) != "none") {
	    	xmlWriter.writeStartElement("border-top");
	    	xmlWriter.writeCharacters(border);
	    	xmlWriter.writeEndElement();	    	
	    }
	      
	    if((border = BORDERS.get(style.getBorderBottom())) != "none") {
	    	xmlWriter.writeStartElement("border-bottom");
	    	xmlWriter.writeCharacters(border);
	    	xmlWriter.writeEndElement();	    	
	    }
	      
	    if((border = BORDERS.get(style.getBorderLeft())) != "none") {
	    	xmlWriter.writeStartElement("border-left");
	    	xmlWriter.writeCharacters(border);
	    	xmlWriter.writeEndElement();	    	
	    }
	      
	    if((border = BORDERS.get(style.getBorderRight())) != "none") {
	    	xmlWriter.writeStartElement("border-right");
	    	xmlWriter.writeCharacters(border);
	    	xmlWriter.writeEndElement();	    	
	    }
	      
	  
	    //Background/fill colour
	    String backgroundColour;
	    if((backgroundColour = helper.getBGColour(style)) != null) {
	    	xmlWriter.writeStartElement("background-color");
	    	xmlWriter.writeCharacters(backgroundColour);
	    	xmlWriter.writeEndElement();	    	
	    }
	    
	    helper.writeFontProperties(xmlWriter, style);
	    
	    xmlWriter.writeEndElement();
	    
	      
  }
    
}
