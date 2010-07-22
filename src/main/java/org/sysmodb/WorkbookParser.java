package org.sysmodb;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.StringWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

public class WorkbookParser {

	private Workbook poi_workbook = null;

	public WorkbookParser(InputStream stream) throws IOException {
		
		File temp = File.createTempFile("poi-data", ".dat");
		OutputStream out=new FileOutputStream(temp);
		temp.deleteOnExit();
		
		byte [] buf = new byte[1024];
		int len;
		while ((len=stream.read(buf))>0) {
			out.write(buf,0,len);
		}
		out.close();		
		try {			
			poi_workbook = new HSSFWorkbook(new FileInputStream(temp));
		} catch (OfficeXmlFileException e) {						
			poi_workbook = new XSSFWorkbook(new FileInputStream(temp));
		}
	}
	
	public String asXML() {
		StringWriter out = new StringWriter();
		XMLWriter writer = new XMLWriter(out, OutputFormat.createPrettyPrint());				
		try {
			writer.write(asXMLDocument());
			writer.close();
		} catch (IOException e) {
			//should never get here, since we are using a StringWriter rather than IO based Writer
			e.printStackTrace();
			return null;
		}
		           
		return out.toString();		
	}
	
	public Document asXMLDocument() {
		Document doc = DocumentHelper.createDocument();
		Element root = doc.addElement("workbook");
		for (int i=0;i<poi_workbook.getNumberOfSheets();i++) {
			Element sheetElement = root.addElement("sheet");
			Sheet sheet = poi_workbook.getSheetAt(i);
			
			sheetElement.addAttribute("name", sheet.getSheetName());
			sheetElement.addAttribute("index", String.valueOf(i+1));
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
					rowElement.addAttribute("index",String.valueOf(y+1));
					int firstCell = row.getFirstCellNum();
					int lastCell = row.getLastCellNum();
					for (int x=firstCell;x<=lastCell;x++) {
						Cell cell = row.getCell(x);
						if (cell !=null) {
							String value=null;
							String type=null;
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK:
								value="";
								type="blank";
								break;
							case Cell.CELL_TYPE_BOOLEAN:
								value=String.valueOf(cell.getBooleanCellValue());
								type="boolean";
								break;
							case Cell.CELL_TYPE_NUMERIC:	
								if (DateUtil.isCellDateFormatted(cell)) {
									type="datetime";
									Date dateCellValue = cell.getDateCellValue();
									SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd'T'H:m:sZ");
									value=format.format(dateCellValue);
								}
								else {
									value=String.valueOf(cell.getNumericCellValue());
									type="numeric";
								}								
								break;
							case Cell.CELL_TYPE_STRING:
								value=cell.getStringCellValue();
								type="string";
								break;
							case Cell.CELL_TYPE_FORMULA:
								type="formula";
								value="=" + cell.getCellFormula();
								break;								
							}
							if (value!=null) {
								Element cellElement = rowElement.addElement("cell");
								cellElement.addAttribute("column",String.valueOf(x+1));
								cellElement.addAttribute("row", String.valueOf(y+1));
								cellElement.addAttribute("type", type);
								cellElement.setText(value);
							}
						}
					}
				}
			}
			
		}
		return doc;
	}
}
