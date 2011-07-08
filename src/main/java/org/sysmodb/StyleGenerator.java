package org.sysmodb;

import org.apache.poi.ss.usermodel.CellStyle;
import java.util.Collections;
import java.util.Map;
import java.util.HashMap;
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