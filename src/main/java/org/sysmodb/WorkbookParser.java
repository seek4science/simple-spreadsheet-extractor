package org.sysmodb;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.LinkedList;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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


/**
 * 
 * @author Stuart Owen, Finn Bacall
 *
 */
public class WorkbookParser {
	
	class CellInfo {
		public String type;
		public String value;
		public String formula;
	}

	private Workbook poi_workbook = null;
	private StyleHelper styleHelper = null;
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
			styleHelper = new HSSFStyleHelper((HSSFWorkbook) poi_workbook);
		} catch (OfficeXmlFileException e) {						
			poi_workbook = new XSSFWorkbook(new FileInputStream(temp));
			styleHelper = new XSSFStyleHelper();
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
	
	public String asCSV(int sheetIndex) {
		return asCSV(sheetIndex,false);
	}
	
	public String asCSV(int sheetIndex, boolean trim) {
		//FIXME: this is a fairly naive implementation, just to get something working quickly.
		//For a better implementation should do something event driven like:
		//https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java
		String result = "";
		Sheet sheet = poi_workbook.getSheetAt(sheetIndex-1);
		int lastRow = sheet.getLastRowNum();
		int lastCol = -1;
		int firstRow = 0;
		int firstCol = 0;
		List<String> stringRowTypes = Arrays.asList(new String[]{"string","datetime"});		
		
		if (trim) {
			firstRow = sheet.getFirstRowNum();
			firstCol = 257;
		}
		
		for (int i=sheet.getFirstRowNum();i<=lastRow;i++) {
			Row row = sheet.getRow(i);		
			if (row !=null && row.getLastCellNum()>lastCol) {
				lastCol=row.getLastCellNum();
			}						
			if (trim && row != null && row.getFirstCellNum()<firstCol ) {
				firstCol = row.getFirstCellNum();
			}
		}
		
		String blankRow = "";
		for (int i=0;i<lastCol-1;i++) {
			blankRow+=",";
		}
		
		for (int y=firstRow;y<=lastRow;y++) {
			Row row = sheet.getRow(y);
			String csvRow = "";
			if (row!=null) {
				for (int x=firstCol;x<lastCol;x++) {
					Cell cell = row.getCell(x);
					CellInfo info = getCellInfo(cell);								
					String value = info.value;
					if (info.type.equalsIgnoreCase("boolean")) value = value.toUpperCase();
					if (stringRowTypes.contains(info.type)) value = "\""+value+"\"";
					csvRow+=value;
					if (x!=lastCol-1) csvRow +=",";				
				}
				result+=csvRow;
			}
			else {
				result+=blankRow;
			}
			
			if (y!=lastRow) result+="\n";
		}
		
		return result;
	}
	
	public Document asXMLDocument() {
		Namespace xmlns = new Namespace("","http://www.sysmo-db.org/2010/xml/spreadsheet");		
		QName workbookName = QName.get("workbook",xmlns);
		
		Document doc = DocumentHelper.createDocument();		
		Element root = doc.addElement(workbookName);

    //Element to hold the cell styles
		Element stylesElement = root.addElement("styles"); 

    //Index: style ID, Value: List of elements (cells) using that style
		ArrayList <LinkedList<Element>> styleMap = new ArrayList<LinkedList<Element>>(poi_workbook.getNumCellStyles());
		for(int i = 0; i < poi_workbook.getNumCellStyles(); i++) {
		  styleMap.add(new LinkedList<Element>()); 
		}
		
		for (int i=0;i<poi_workbook.getNumberOfSheets();i++) {			
			Element sheetElement = root.addElement("sheet");			
			
			Sheet sheet = poi_workbook.getSheetAt(i);
			
			sheetElement.addAttribute("name", sheet.getSheetName());
			sheetElement.addAttribute("index", String.valueOf(i+1));
			sheetElement.addAttribute("hidden", String.valueOf(poi_workbook.isSheetHidden(i)));
			sheetElement.addAttribute("very_hidden", String.valueOf(poi_workbook.isSheetVeryHidden(i)));

			//Columns (for column widths - styling)
			int lastCol=1;
			int firstCol=1;
			Element columnsElement = sheetElement.addElement("columns");
			
			int firstRow=sheet.getFirstRowNum();
			int lastRow=sheet.getLastRowNum();
			
			Element rowsElement = sheetElement.addElement("rows");
			rowsElement.addAttribute("first_row", String.valueOf(firstRow+1));
			rowsElement.addAttribute("last_row", String.valueOf(lastRow+1));
			
			for (int y=firstRow;y<=lastRow;y++) {
				Row row = sheet.getRow(y);
				if (row!=null) {
					Element rowElement = rowsElement.addElement("row");					
					rowElement.addAttribute("index",String.valueOf(y+1));
					//Get height of row, if different from default
					if(sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints())
					    rowElement.addAttribute("height",""+row.getHeightInPoints()+"pt");
					
					int firstCell = row.getFirstCellNum();
					
					//Skip if row has no cells
					if(firstCell == -1)
            continue;
					
					if(firstCell < firstCol)
				    firstCol = firstCell + 1;//Number of columns
					
					int lastCell = row.getLastCellNum();					
					if(lastCell > lastCol)
				    lastCol = lastCell;//Number of columns	
					
					for (int x=firstCell;x<=lastCell;x++) {
						Cell cell = row.getCell(x);
						if (cell !=null) {
							CellInfo info = getCellInfo(cell);
							
							if (info.value!=null) {								
								Element cellElement = rowElement.addElement("cell");								
								cellElement.addAttribute("column",String.valueOf(x+1));
								cellElement.addAttribute("column_alpha",column_alpha(x));
								cellElement.addAttribute("row", String.valueOf(y+1));
								cellElement.addAttribute("type", info.type);

								//Cell style
								//Add to style linked list
								int styleIndex = cell.getCellStyle().getIndex();
								
                styleMap.get(styleIndex).add(cellElement);
								cellElement.addAttribute("style", ("style"+cell.getCellStyle().getIndex()));
								
								if (info.formula!=null) {
									cellElement.addAttribute("formula", info.formula);
								}
								cellElement.setText(info.value);
							}
						}
					}
				}
			}		
			
  		columnsElement.addAttribute("first_column", String.valueOf(firstCol));
			columnsElement.addAttribute("last_column", String.valueOf(lastCol));
			for(int x = firstCol-1; x < lastCol; x++)
			{
				Element columnElement = columnsElement.addElement("column");
				columnElement.addAttribute("index", String.valueOf(x+1));
				columnElement.addAttribute("column_alpha", String.valueOf(column_alpha(x)));
				columnElement.addAttribute("width", String.valueOf(sheet.getColumnWidth(x)));
			}			
		}
		
  	//Remove duplicate styles
    Hashtable<Integer, Short> styleHashTable = new Hashtable<Integer, Short>();
    
    //Add style info to style element
    for (short s=0;s<poi_workbook.getNumCellStyles();s++) {
      
      //Don't bother rendering styles that aren't used in any cells!
      if(styleMap.get(s).isEmpty())
        continue;
      
      CellStyle style;
      try
      {
        style = poi_workbook.getCellStyleAt(s);
      }
      //Sometimes XSLX messes up and reports wrong number of
      // styles...
      catch(IndexOutOfBoundsException e)
      {
        break;
      }

      Element styleElement = stylesElement.addElement("style");
      styleElement.addAttribute("id",("style"+s));
      StyleGenerator.createStyle(style,styleElement,styleHelper);

      //If we have an empty style element, its useless, so don't display it
      if(styleElement.elements().isEmpty())
      {
        styleElement.detach();
        //Remove "style" attributes from cells that were linked to the blank style
        Iterator <Element> iter = styleMap.get(s).iterator();
        while(iter.hasNext())
        {
          iter.next().addAttribute("style",null);
        }
      }
      else
      {
        int styleHash = StyleGenerator.getStyleHash(styleElement);
        //Check a duplicate style doesn't already exist, if it does, point all cells to that style and delete
        // the duplicate
        if(styleHashTable.containsKey(styleHash))
        {            
          styleElement.detach();
          Iterator <Element> iter = styleMap.get(s).iterator();
          while(iter.hasNext())
          {
            iter.next().addAttribute("style","style"+styleHashTable.get(styleHash).toString());
          }
        }
        else
        {
          styleHashTable.put(styleHash, s);        
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
	
	private CellInfo getCellInfo(Cell cell) {
		//FIXME: too long and duplicates, needs refactoring
		
		CellInfo info = new CellInfo();
		if (cell==null) {
			info.value="";
			info.type="blank";
			info.formula=null;
		}
		else {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				info.value="";
				info.type="blank";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				info.value=String.valueOf(cell.getBooleanCellValue());
				info.type="boolean";
				break;
			case Cell.CELL_TYPE_NUMERIC:	
				if (DateUtil.isCellDateFormatted(cell)) {
					info.type="datetime";
					Date dateCellValue = cell.getDateCellValue();									
					info.value=dateFormatter.format(dateCellValue);
				}
				else {
					info.value=String.valueOf(cell.getNumericCellValue());
					info.type="numeric";
				}								
				break;
			case Cell.CELL_TYPE_STRING:
				info.value=cell.getStringCellValue();
				info.type="string";
				break;
			case Cell.CELL_TYPE_FORMULA:								
				
				FormulaEvaluator evaluator = poi_workbook.getCreationHelper().createFormulaEvaluator();
				CellValue cellValue = evaluator.evaluate(cell);								
				info.value=cellValue.formatAsString();
				info.formula=cell.getCellFormula();
				switch(cellValue.getCellType()) {
				case Cell.CELL_TYPE_BOOLEAN:
					info.type="boolean";
					break;
				case Cell.CELL_TYPE_STRING:
					info.type="string";
					break;
				case Cell.CELL_TYPE_NUMERIC:											
					info.type="numeric";					
					if (DateUtil.isCellDateFormatted(cell)) {
						info.type="datetime";																		
						info.value=dateFormatter.format(DateUtil.getJavaDate(cellValue.getNumberValue()));
					}
					break;								
				}
											
				break;								
			}
		}
		return info;
	}
}
