/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.ss.usermodel.CellStyle;
import org.dom4j.Element;

/**
 *
 * @author Finn
 */
public abstract class StyleGenerator {
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

  public static boolean isStyleEmpty(CellStyle style,StyleHelper helper) {
	  
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
  
  public static void writeStyle(XMLStreamWriter xmlWriter, CellStyle style, StyleHelper helper) throws XMLStreamException {
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
  
  public static void createStyle(CellStyle style, Element element, StyleHelper helper)
  {
    //Border properties
    String border = "none";
    if((border = BORDERS.get(style.getBorderTop())) != "none")
      element.addElement("border-top").addText(border);
    if((border = BORDERS.get(style.getBorderBottom())) != "none")
      element.addElement("border-bottom").addText(border);
    if((border = BORDERS.get(style.getBorderLeft())) != "none")
      element.addElement("border-left").addText(border);
    if((border = BORDERS.get(style.getBorderRight())) != "none")
      element.addElement("border-right").addText(border);
  
    //Background/fill colour
    String backgroundColour;
    if((backgroundColour = helper.getBGColour(style)) != null)
      element.addElement("background-color").addText(backgroundColour);
    
    //Font properties
    helper.setFontProperties(style, element);
  }
  
  public static int getStyleHash(Element element)
  {
    return element.asXML().replaceAll("style([0-9]+)","").hashCode();
  }
}
