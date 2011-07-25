package org.sysmodb;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.dom4j.Element;


/**
 *
 * @author Finn
 */
public class XSSFStyleHelper implements StyleHelper{

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
      element.addElement("font-size").addText(String.valueOf(font.getFontHeightInPoints()) + "pt");
    if(!font.getFontName().equals(XSSFFont.DEFAULT_FONT_NAME))
      element.addElement("font-family").addText(font.getFontName());
    if(font.getColor() != XSSFFont.DEFAULT_FONT_COLOR)
    {
      String colorString = getRGBString(font.getXSSFColor());
      if(colorString != null)
      {
        element.addElement("color").addText(colorString);  
      }      
    }
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
      String rgb = colour.getARGBHex();
      //XSSF has a bug where the above can sometimes return null
      // so we check here
      if(rgb != null)
      {        
        if(rgb.length() > 6)
          rgb = rgb.substring(2,rgb.length());
        string = "#" + rgb;
      }
    }

    return string;
  }
}
