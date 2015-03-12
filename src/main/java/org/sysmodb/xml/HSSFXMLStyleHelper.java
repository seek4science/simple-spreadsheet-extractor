/*******************************************************************************
 * Copyright (c) 2009-2013, University of Manchester
 *  
 * Licensed under the New BSD License. 
 * Please see LICENSE file that is distributed with the source code
 ******************************************************************************/
package org.sysmodb.xml;

import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 
 * @author Finn, Stuart Owen
 */

public class HSSFXMLStyleHelper implements XMLStyleHelper {

	private static HSSFWorkbook workbook;
	private static HSSFPalette palette;

	private static final HSSFColor HSSF_AUTO = new HSSFColor.AUTOMATIC();

	public HSSFXMLStyleHelper(HSSFWorkbook wb) {
		workbook = wb;
		palette = wb.getCustomPalette();
	}

	public String getBGColour(CellStyle style) {
		return getRGBString(style.getFillForegroundColor());
	}

	public boolean areFontsEmpty(CellStyle style) {
		HSSFCellStyle newStyle = (HSSFCellStyle) style;
		HSSFFont font = newStyle.getFont(workbook);
		if (font.getBoldweight() == HSSFFont.BOLDWEIGHT_BOLD)
			return false;
		if (font.getItalic())
			return false;
		if (font.getUnderline() != HSSFFont.U_NONE)
			return false;
		// Ignore same-ish defaults
		if (font.getFontHeightInPoints() != 10
				&& font.getFontHeightInPoints() != 11)
			return false;
		// Arial is default for Excel, Calibri is default for OO
		if (!font.getFontName().equals("Arial")
				&& !font.getFontName().equals("Calibri"))
			return false;
		if ((font.getColor() != HSSFFont.COLOR_NORMAL)
				&& (getRGBString(font.getColor()) != null)
				&& !getRGBString(font.getColor()).equals("#000"))
			return false;

		return true;
	}

	@Override
	public void writeFontProperties(XMLStreamWriter xmlWriter, CellStyle style)
			throws XMLStreamException {
		HSSFCellStyle newStyle = (HSSFCellStyle) style;
		HSSFFont font = newStyle.getFont(workbook);
		if (font.getBoldweight() == HSSFFont.BOLDWEIGHT_BOLD) {
			xmlWriter.writeStartElement("font-weight");
			xmlWriter.writeCharacters("bold");
			xmlWriter.writeEndElement();
		}
		if (font.getItalic()) {
			xmlWriter.writeStartElement("font-style");
			xmlWriter.writeCharacters("italics");
			xmlWriter.writeEndElement();
		}
		if (font.getUnderline() != HSSFFont.U_NONE) {
			xmlWriter.writeStartElement("text-decoration");
			xmlWriter.writeCharacters("underline");
			xmlWriter.writeEndElement();
		}
		// Ignore same-ish defaults
		if (font.getFontHeightInPoints() != 10
				&& font.getFontHeightInPoints() != 11) {
			xmlWriter.writeStartElement("font-size");
			xmlWriter.writeCharacters(String.valueOf(font
					.getFontHeightInPoints() + "pt"));
			xmlWriter.writeEndElement();
		}
		// Arial is default for Excel, Calibri is default for OO
		if (!font.getFontName().equals("Arial")
				&& !font.getFontName().equals("Calibri")) {
			xmlWriter.writeStartElement("font-family");
			xmlWriter.writeCharacters(font.getFontName());
			xmlWriter.writeEndElement();
		}
		if ((font.getColor() != HSSFFont.COLOR_NORMAL)
				&& (getRGBString(font.getColor()) != null)
				&& !getRGBString(font.getColor()).equals("#000")) {
			xmlWriter.writeStartElement("color");
			xmlWriter.writeCharacters(getRGBString(font.getColor()));
			xmlWriter.writeEndElement();
		}

	}

	private String getRGBString(short index) {
		String string = null;
		HSSFColor color = palette.getColor(index);

		if (index == HSSF_AUTO.getIndex() || color == null) {

		} else {
			short[] rgb = color.getTriplet();
			string = "#";
			for (int i = 0; i <= 2; i++) {
				String colourSection = Integer.toHexString((int) rgb[i]);
				if (colourSection.length() == 1)
					colourSection = "0" + colourSection;

				string += colourSection;
			}
		}

		return string;
	}

}
