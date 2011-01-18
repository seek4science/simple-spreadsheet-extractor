package org.sysmodb;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Element;


/**
 *
 * @author Finn
 */
public class XSSFStyleHelper implements StyleHelper{

    private static XSSFWorkbook workbook;

    public XSSFStyleHelper(XSSFWorkbook wb)
    {
	workbook = wb;
    }

    public String getBGColour(CellStyle style)
    {
	XSSFCellStyle newStyle = (XSSFCellStyle) style;
	XSSFColor colour = newStyle.getFillForegroundXSSFColor();
	return getRGBString(colour);
    }

    public void setFontProperties(CellStyle style, Element element)
    {
	XSSFCellStyle newStyle = (XSSFCellStyle) style;
	XSSFFont font = newStyle.getFont();
	if(font.getBold())
	    element.addElement("font-weight").addText("bold");
	if(font.getItalic())
	    element.addElement("font-style").addText("italics");
	if(font.getUnderline() != XSSFFont.U_NONE)
	    element.addElement("text-decoration").addText("underline");
	if(font.getFontHeight() != XSSFFont.DEFAULT_FONT_SIZE)
	    element.addElement("font-size").addText(String.valueOf(font.getFontHeight() + "pt"));
	if(!font.getFontName().equals(XSSFFont.DEFAULT_FONT_NAME))
	    element.addElement("font-family").addText(font.getFontName());
	if(font.getColor() != XSSFFont.DEFAULT_FONT_COLOR)
	    element.addElement("color").addText(getRGBString(font.getXSSFColor()));
    }


    private String getRGBString(XSSFColor colour)
    {
	String string = null;
	//Disregard default/automatic colours, to avoid cluttering XML
        if (colour == null || colour.isAuto())
	{
	    return string;
	}
	else
	{
	    byte[] rgb = colour.getRgb();
	    string = "#";
	    for(int i = 0; i <= 2; i++)
	    {
	    	String colourSection = Integer.toHexString((int) rgb[i]);
	    	if(colourSection.length() == 1)
	    		colourSection = "0" + colourSection;
	    	
	    	string += colourSection;
	    }
	}

	return string;
    }
}
