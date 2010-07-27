package org.sysmodb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.Namespace;
import org.dom4j.QName;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

public class WorkbookParser {

	private Workbook poi_workbook = null;
	private final static SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'H:m:sZ");

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
		Namespace xmlns = new Namespace("","http://www.sysmo-db.org/2010/xml/spreadsheet");		
		QName workbookName = QName.get("workbook",xmlns);
		
		Document doc = DocumentHelper.createDocument();		
		Element root = doc.addElement(workbookName);
		
		for (int i=0;i<poi_workbook.getNumberOfSheets();i++) {			
			Element sheetElement = root.addElement("sheet");			
			
			Sheet sheet = poi_workbook.getSheetAt(i);
			
			sheetElement.addAttribute("name", sheet.getSheetName());
			sheetElement.addAttribute("index", String.valueOf(i+1));
			sheetElement.addAttribute("hidden", String.valueOf(poi_workbook.isSheetHidden(i)));
			sheetElement.addAttribute("very_hidden", String.valueOf(poi_workbook.isSheetVeryHidden(i)));
			
			int firstRow=sheet.getFirstRowNum();
			int lastRow=sheet.getLastRowNum();
			sheetElement.addAttribute("first_row", String.valueOf(firstRow+1));
			sheetElement.addAttribute("last_row", String.valueOf(lastRow+1));			
			for (int y=firstRow;y<=lastRow;y++) {
				Row row = sheet.getRow(y);
				if (row!=null) {
					Element rowElement = sheetElement.addElement("row");					
					rowElement.addAttribute("index",String.valueOf(y+1));
					int firstCell = row.getFirstCellNum();
					int lastCell = row.getLastCellNum();
					String formula=null;
					for (int x=firstCell;x<=lastCell;x++) {
						Cell cell = row.getCell(x);
						if (cell !=null) {
							//FIXME: too long and duplicates, needs refactoring
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
									value=dateFormatter.format(dateCellValue);
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
								
								FormulaEvaluator evaluator = poi_workbook.getCreationHelper().createFormulaEvaluator();
								CellValue cellValue = evaluator.evaluate(cell);								
								value=cellValue.formatAsString();
								formula=cell.getCellFormula();
								switch(cellValue.getCellType()) {
								case Cell.CELL_TYPE_BOOLEAN:
									type="boolean";
									break;
								case Cell.CELL_TYPE_STRING:
									type="string";
									break;
								case Cell.CELL_TYPE_NUMERIC:											
									type="numeric";					
									if (DateUtil.isCellDateFormatted(cell)) {
										type="datetime";																		
										value=dateFormatter.format(DateUtil.getJavaDate(cellValue.getNumberValue()));
									}
									break;								
								}
															
								break;								
							}
							if (value!=null) {								
								Element cellElement = rowElement.addElement("cell");								
								cellElement.addAttribute("column",String.valueOf(x+1));
								cellElement.addAttribute("column_alpha",column_alpha(x));
								cellElement.addAttribute("row", String.valueOf(y+1));
								cellElement.addAttribute("type", type);
								if (formula!=null) {
									cellElement.addAttribute("formula", formula);
								}
								cellElement.setText(value);
							}
						}
					}
				}
			}
			
		}
		return doc;
	}

	private String column_alpha(int col) {
		String result = "";
		while (col>-1) {
			int letter = (col % 26);
			result = Character.toString((char)(letter+65)) + result;
			col = (col / 26) - 1;
		}
		return result;
	}
}
