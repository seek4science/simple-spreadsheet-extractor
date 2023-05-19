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

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 
 * @author Finn
 */
public abstract class XMLStyleGenerator {
	private static final Map<BorderStyle, String> BORDERS = createBorderMap();

	private static Map<BorderStyle, String> createBorderMap() {
		Map<BorderStyle, String> result = new HashMap<BorderStyle, String>();
		result.put(BorderStyle.DASHED, "dashed 1pt");
		result.put(BorderStyle.DASH_DOT, "dashed 1pt");
		result.put(BorderStyle.DASH_DOT_DOT, "dashed 1pt");
		result.put(BorderStyle.DOTTED, "dotted 1pt");
		result.put(BorderStyle.DOUBLE, "double 3pt");
		result.put(BorderStyle.HAIR, "solid 1pt");
		result.put(BorderStyle.MEDIUM, "2pt solid");
		result.put(BorderStyle.MEDIUM_DASHED, "2pt dashed");
		result.put(BorderStyle.MEDIUM_DASH_DOT, "2pt dashed");
		result.put(BorderStyle.MEDIUM_DASH_DOT_DOT, "2pt dashed");
		result.put(BorderStyle.NONE, "none");
		result.put(BorderStyle.SLANTED_DASH_DOT, "dashed 2pt");
		result.put(BorderStyle.THICK, "solid 3pt");
		result.put(BorderStyle.THIN, "dashed 1pt");

		return Collections.unmodifiableMap(result);
	}

	public static boolean isStyleEmpty(CellStyle style, XMLStyleHelper helper) {

		if (style.getBorderTop() != BorderStyle.NONE)
			return false;
		if (style.getBorderBottom() != BorderStyle.NONE)
			return false;
		if (style.getBorderLeft() != BorderStyle.NONE)
			return false;
		if (style.getBorderRight() != BorderStyle.NONE)
			return false;

		// Background/fill colour
		if ((helper.getBGColour(style)) != null)
			return false;

		return helper.areFontsEmpty(style);
	}

	public static void writeStyle(XMLStreamWriter xmlWriter, CellStyle style, XMLStyleHelper helper)
			throws XMLStreamException {
		String border = "none";
		xmlWriter.writeStartElement("style");
		xmlWriter.writeAttribute("id", "style" + style.getIndex());
		if ((border = BORDERS.get(style.getBorderTop())) != "none") {
			xmlWriter.writeStartElement("border-top");
			xmlWriter.writeCharacters(border);
			xmlWriter.writeEndElement();
		}

		if ((border = BORDERS.get(style.getBorderBottom())) != "none") {
			xmlWriter.writeStartElement("border-bottom");
			xmlWriter.writeCharacters(border);
			xmlWriter.writeEndElement();
		}

		if ((border = BORDERS.get(style.getBorderLeft())) != "none") {
			xmlWriter.writeStartElement("border-left");
			xmlWriter.writeCharacters(border);
			xmlWriter.writeEndElement();
		}

		if ((border = BORDERS.get(style.getBorderRight())) != "none") {
			xmlWriter.writeStartElement("border-right");
			xmlWriter.writeCharacters(border);
			xmlWriter.writeEndElement();
		}

		// Background/fill colour
		String backgroundColour;
		if ((backgroundColour = helper.getBGColour(style)) != null) {
			xmlWriter.writeStartElement("background-color");
			xmlWriter.writeCharacters(backgroundColour);
			xmlWriter.writeEndElement();
		}

		helper.writeFontProperties(xmlWriter, style);

		xmlWriter.writeEndElement();

	}

}
