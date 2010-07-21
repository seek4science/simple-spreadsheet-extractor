package org.sysmodb;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;

public class Workbook {
	
	private org.apache.poi.ss.usermodel.Workbook poi_workbook = null;
	
	public Workbook(InputStream stream) throws IOException {
		try {
			poi_workbook = new HSSFWorkbook(new BufferedInputStream(stream));
		}
		catch(OfficeXmlFileException e) { //
			//TODO: handle XSLX documents here
			throw e;
		}
	}
}
