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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * 
 * @author Finn
 */
public class XSSFXMLStyleHelper implements XMLStyleHelper {

	public String getBGColour(CellStyle style) {
		XSSFCellStyle newStyle = (XSSFCellStyle) style;
		XSSFColor colour = newStyle.getFillForegroundXSSFColor();
		return getRGBString(colour);
	}

	public boolean areFontsEmpty(CellStyle style) {
		XSSFCellStyle newStyle = (XSSFCellStyle) style;
		XSSFFont font = newStyle.getFont();
		if (font.getBold())
			return false;
		if (font.getItalic())
			return false;
		if (font.getUnderline() != XSSFFont.U_NONE)
			return false;
		if (font.getFontHeight() != XSSFFont.DEFAULT_FONT_SIZE)
			return false;
		if (!font.getFontName().equals(XSSFFont.DEFAULT_FONT_NAME))
			return false;
		if (font.getColor() != XSSFFont.DEFAULT_FONT_COLOR) {
			String colorString = getRGBString(font.getXSSFColor());
			if (colorString != null) {
				return false;
			}
		}
		return true;
	}

	@Override
	public void writeFontProperties(XMLStreamWriter xmlWriter, CellStyle style)
			throws XMLStreamException {
		XSSFCellStyle newStyle = (XSSFCellStyle) style;
		XSSFFont font = newStyle.getFont();
		if (font.getBold()) {
			xmlWriter.writeStartElement("font-weight");
			xmlWriter.writeCharacters("bold");
			xmlWriter.writeEndElement();
		}
		if (font.getItalic()) {
			xmlWriter.writeStartElement("font-style");
			xmlWriter.writeCharacters("italics");
			xmlWriter.writeEndElement();
		}
		if (font.getUnderline() != XSSFFont.U_NONE) {
			xmlWriter.writeStartElement("text-decoration");
			xmlWriter.writeCharacters("underline");
			xmlWriter.writeEndElement();
		}
		if (font.getFontHeight() != XSSFFont.DEFAULT_FONT_SIZE) {
			xmlWriter.writeStartElement("font-size");
			xmlWriter.writeCharacters(String.valueOf(font
					.getFontHeightInPoints()) + "pt");
			xmlWriter.writeEndElement();
		}

		if (!font.getFontName().equals(XSSFFont.DEFAULT_FONT_NAME)) {
			xmlWriter.writeStartElement("font-family");
			xmlWriter.writeCharacters(String.valueOf(font.getFontName()));
			xmlWriter.writeEndElement();
		}
		if (font.getColor() != XSSFFont.DEFAULT_FONT_COLOR) {
			String colorString = getRGBString(font.getXSSFColor());
			if (colorString != null) {
				xmlWriter.writeStartElement("color");
				xmlWriter.writeCharacters(colorString);
				xmlWriter.writeEndElement();
			}
		}

	}

	private String getRGBString(XSSFColor colour) {
		String string = null;
		// Disregard default/automatic colours, to avoid cluttering XML
		if (colour == null || colour.isAuto()) {
			return string;
		} else {
			String rgb = colour.getARGBHex();
			// XSSF has a bug where the above can sometimes return null
			// so we check here
			if (rgb != null) {
				if (rgb.length() > 6)
					rgb = rgb.substring(2, rgb.length());
				string = "#" + rgb;
			}
		}

		return string;
	}

}
