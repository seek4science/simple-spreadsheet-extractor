package org.sysmodb;

import java.io.IOException;
import java.io.Writer;

import org.dom4j.Attribute;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

/**
 * Specialised writer that strips out on control characters illegal for XML v1.0 
 * These characters are simply removed completely
 * 
 * @author Stuart Owen
 *
 */
public class ControlCharStrippingXMLWriter extends XMLWriter {
	
	public ControlCharStrippingXMLWriter(Writer writer, OutputFormat format) {
		super(writer, format);
	}
	
	@Override
	protected void writeAttributes(Element element) throws IOException { 
		for (Object at : element.attributes()) {
			((Attribute)at).setText(stripControlCharacters(((Attribute)at).getText()));
		}
		Attribute attribute = element.attribute("formula");
		if (attribute!=null) {
			attribute.setText(stripControlCharacters(attribute.getText()));				
		}	
		if (element.getText()!=null && element.getText().length()>0) {
			element.setText(stripControlCharacters(element.getText()));
		}
		super.writeAttributes(element);
	}		
	
	private String stripControlCharacters(String original) {
		return original.replaceAll("\\p{Cntrl}", "");
	}
}
