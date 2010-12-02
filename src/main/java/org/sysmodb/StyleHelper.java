package org.sysmodb;

import org.apache.poi.ss.usermodel.CellStyle;
import org.dom4j.Element;

 /**
 *
 * @author Finn
 */
public interface StyleHelper {

    String getBGColour(CellStyle style);

    void setFontProperties(CellStyle style, Element element);
}
