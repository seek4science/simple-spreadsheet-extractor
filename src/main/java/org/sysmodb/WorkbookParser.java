package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;

public class WorkbookParser {

	private Workbook poi_workbook = null;

	public WorkbookParser(InputStream stream) throws IOException {
		try {
			poi_workbook = new HSSFWorkbook(stream);
		} catch (OfficeXmlFileException e) { //
			poi_workbook = new XSSFWorkbook(stream);
		}
	}
	
	public String asXML() {
		return asXMLDocument().asXML();
		
		
	}
	
	public Document asXMLDocument() {
		Document doc = DocumentHelper.createDocument();
		Element root = doc.addElement("workbook");
		for (int i=0;i<poi_workbook.getNumberOfSheets();i++) {
			Element sheetElement = root.addElement("sheet");
			Sheet sheet = poi_workbook.getSheetAt(i);
			
			sheetElement.addAttribute("name", sheet.getSheetName());
			sheetElement.addAttribute("index", String.valueOf(i));
			sheetElement.addAttribute("hidden", String.valueOf(poi_workbook.isSheetHidden(i)));
			sheetElement.addAttribute("very_hidden", String.valueOf(poi_workbook.isSheetVeryHidden(i)));
			
			int firstRow=sheet.getFirstRowNum();
			int lastRow=sheet.getLastRowNum();
			sheetElement.addAttribute("first_row", String.valueOf(firstRow));
			sheetElement.addAttribute("last_row", String.valueOf(lastRow));			
			for (int y=firstRow;y<=lastRow;y++) {
				Row row = sheet.getRow(y);
				if (row!=null) {
					Element rowElement = sheetElement.addElement("row");
					rowElement.addAttribute("index",String.valueOf(y));
					int firstCell = row.getFirstCellNum();
					int lastCell = row.getLastCellNum();
					for (int x=firstCell;x<=lastCell;x++) {
						Cell cell = row.getCell(x);
						if (cell !=null) {
							String value=null;
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK:
								value="";
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								value=String.valueOf(cell.getBooleanCellValue());
								break;
							case Cell.CELL_TYPE_NUMERIC:
								value=String.valueOf(cell.getNumericCellValue());
								break;
							case Cell.CELL_TYPE_STRING:
								value=cell.getStringCellValue();
								break;
							}
							Element cellElement = rowElement.addElement("cell");
							cellElement.addAttribute("column",String.valueOf(x));
							cellElement.addAttribute("row", String.valueOf(y));
							cellElement.setText(value);						
						}
					}
				}
			}
			
		}
		return doc;
	}
}
