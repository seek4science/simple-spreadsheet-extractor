package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Workbook {

	private org.apache.poi.ss.usermodel.Workbook poi_workbook = null;

	public Workbook(InputStream stream) throws IOException {
		try {
			poi_workbook = new HSSFWorkbook(stream);
		} catch (OfficeXmlFileException e) { //
			poi_workbook = new XSSFWorkbook(stream);
		}
	}
}
