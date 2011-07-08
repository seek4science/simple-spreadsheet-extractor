package org.sysmodb;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.dom4j.Element;

/**
 *
 * @author Finn
 */

public class HSSFStyleHelper implements StyleHelper {

  private static HSSFWorkbook workbook;
  private static HSSFPalette palette;

  private static final HSSFColor HSSF_AUTO = new HSSFColor.AUTOMATIC();


  public HSSFStyleHelper(HSSFWorkbook wb)
  {
    workbook = wb;
    palette = wb.getCustomPalette();
  }

  public String getBGColour(CellStyle style)
  {
    return getRGBString(style.getFillForegroundColor());
  }

  public void setFontProperties(CellStyle style, Element element)
  {
    HSSFCellStyle newStyle = (HSSFCellStyle) style;
    HSSFFont font = newStyle.getFont(workbook);
    if(font.getBoldweight() == HSSFFont.BOLDWEIGHT_BOLD)
      element.addElement("font-weight").addText("bold");
    if(font.getItalic())
      element.addElement("font-style").addText("italics");
    if(font.getUnderline() != HSSFFont.U_NONE)
      element.addElement("text-decoration").addText("underline");
    //Ignore same-ish defaults
    if(font.getFontHeightInPoints() != 10 && font.getFontHeightInPoints() != 11)
      element.addElement("font-size").addText(String.valueOf(font.getFontHeightInPoints() + "pt"));
    //Arial is default for Excel, Calibri is default for OO
    if(!font.getFontName().equals("Arial") && !font.getFontName().equals("Calibri"))
      element.addElement("font-family").addText(font.getFontName());
    if((font.getColor() != HSSFFont.COLOR_NORMAL) && !getRGBString(font.getColor()).equals("#000"))
      element.addElement("color").addText(getRGBString(font.getColor()));
  }

  private String getRGBString(short index)
  {
    String string = null;
    HSSFColor color = palette.getColor(index);

    if (index == HSSF_AUTO.getIndex() || color == null)
    {
    
    }
    else
    {
      short[] rgb = color.getTriplet();
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
