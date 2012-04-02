package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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

	private Workbook poiWorkbook = null;
	private StyleHelper styleHelper = null;
	
	public WorkbookParser(InputStream stream) throws IOException,
			InvalidFormatException {

		poiWorkbook = WorkbookFactory.create(stream);

		if (poiWorkbook instanceof XSSFWorkbook) {
			styleHelper = new XSSFStyleHelper();
		} else {
			styleHelper = new HSSFStyleHelper((HSSFWorkbook) poiWorkbook);			
		}
	}

	public String asXML() {
		StringWriter out = new StringWriter();
		OutputFormat format = OutputFormat.createPrettyPrint();
		format.setEncoding("UTF-8");		
		XMLWriter writer = new ControlCharStrippingXMLWriter(out, format);
		
		writer.setEscapeText(true);						
		
		try {			
			writer.write(asXMLDocument());
			writer.close();
		} catch (IOException e) {
			// should never get here, since we are using a StringWriter rather
			// than IO based Writer
			e.printStackTrace();
			return null;
		}

		return out.toString();
	}

	public String asCSV(int sheetIndex) {
		return asCSV(sheetIndex, false);
	}

	public String asCSV(int sheetIndex, boolean trim) {
		// FIXME: this is a fairly naive implementation, just to get something
		// working quickly.
		// For a better implementation should do something event driven like:
		// https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/hssf/eventusermodel/examples/XLS2CSVmra.java
		String result = "";
		Sheet sheet = poiWorkbook.getSheetAt(sheetIndex - 1);
		int lastRow = sheet.getLastRowNum();
		int lastCol = -1;
		int firstRow = 0;
		int firstCol = 0;
		List<String> stringRowTypes = Arrays.asList(new String[] { "string",
				"datetime" });

		if (trim) {
			firstRow = sheet.getFirstRowNum();
			firstCol = 257;
		}

		for (int i = sheet.getFirstRowNum(); i <= lastRow; i++) {
			Row row = sheet.getRow(i);
			if (row != null && row.getLastCellNum() > lastCol) {
				lastCol = row.getLastCellNum();
			}
			if (trim && row != null && row.getFirstCellNum() < firstCol) {
				firstCol = row.getFirstCellNum();
			}
		}

		String blankRow = "";
		for (int i = 0; i < lastCol - 1; i++) {
			blankRow += ",";
		}

		for (int y = firstRow; y <= lastRow; y++) {
			Row row = sheet.getRow(y);
			String csvRow = "";
			if (row != null) {
				for (int x = firstCol; x < lastCol; x++) {
					Cell cell = row.getCell(x);
					CellInfo info = new CellInfo(cell,poiWorkbook);
					String value = info.value;
					if (info.type.equalsIgnoreCase("boolean"))
						value = value.toUpperCase();
					if (stringRowTypes.contains(info.type))
						value = "\"" + value + "\"";
					csvRow += value;
					if (x != lastCol - 1)
						csvRow += ",";
				}
				result += csvRow;
			} else {
				result += blankRow;
			}

			if (y != lastRow)
				result += "\n";
		}

		return result;
	}

	public Document asXMLDocument() {
		Namespace xmlns = new Namespace("",
				"http://www.sysmo-db.org/2010/xml/spreadsheet");
		QName workbookName = QName.get("workbook", xmlns);
		
		Document doc = DocumentHelper.createDocument();		
		
		Element root = doc.addElement(workbookName);
		
		Element namedRangesElement = root.addElement("named_ranges");
		namedRangesToXML(namedRangesElement);

		// Element to hold the cell styles
		Element stylesElement = root.addElement("styles");

		// Index: style ID, Value: List of elements (cells) using that style
		ArrayList<LinkedList<Element>> styleMap = new ArrayList<LinkedList<Element>>(
				poiWorkbook.getNumCellStyles());
		for (int i = 0; i < poiWorkbook.getNumCellStyles(); i++) {
			styleMap.add(new LinkedList<Element>());
		}

		for (int i = 0; i < poiWorkbook.getNumberOfSheets(); i++) {
			Element sheetElement = root.addElement("sheet");

			Sheet sheet = poiWorkbook.getSheetAt(i);

			sheetToXML(styleMap, i, sheetElement, sheet);
		}

		// Remove duplicate styles
		Hashtable<Integer, Short> styleHashTable = new Hashtable<Integer, Short>();

		// Add style info to style element
		for (short s = 0; s < poiWorkbook.getNumCellStyles(); s++) {

			// Don't bother rendering styles that aren't used in any cells!
			if (styleMap.get(s).isEmpty())
				continue;

			CellStyle style;
			try {
				style = poiWorkbook.getCellStyleAt(s);
			}
			// Sometimes XSLX messes up and reports wrong number of
			// styles...
			catch (IndexOutOfBoundsException e) {
				break;
			}

			Element styleElement = stylesElement.addElement("style");
			styleElement.addAttribute("id", ("style" + s));
			StyleGenerator.createStyle(style, styleElement, styleHelper);

			// If we have an empty style element, its useless, so don't display
			// it
			if (styleElement.elements().isEmpty()) {
				styleElement.detach();
				// Remove "style" attributes from cells that were linked to the
				// blank style
				Iterator<Element> iter = styleMap.get(s).iterator();
				while (iter.hasNext()) {
					iter.next().addAttribute("style", null);
				}
			} else {
				int styleHash = StyleGenerator.getStyleHash(styleElement);
				// Check a duplicate style doesn't already exist, if it does,
				// point all cells to that style and delete
				// the duplicate
				if (styleHashTable.containsKey(styleHash)) {
					styleElement.detach();
					Iterator<Element> iter = styleMap.get(s).iterator();
					while (iter.hasNext()) {
						iter.next().addAttribute(
								"style",
								"style"
										+ styleHashTable.get(styleHash)
												.toString());
					}
				} else {
					styleHashTable.put(styleHash, s);
				}
			}
		}

		return doc;
	}

	private void namedRangesToXML(Element namedRangesElement) {
		for(int i = 0; i < poiWorkbook.getNumberOfNames(); i++) {
            Name name = poiWorkbook.getNameAt(i);            
            try {
            	if(!name.isDeleted() && !name.isFunctionName()) {                	
                	String formula = name.getRefersToFormula();                	
                	AreaReference areaReference = new AreaReference(formula);
                    CellReference firstCellReference = areaReference.getFirstCell();
                    CellReference lastCellReference = areaReference.getLastCell();
                    formula = formula.replaceAll("\\p{C}", "?");
                    
                    Element namedRangeElement = namedRangesElement.addElement("named_range");
                    namedRangeElement.addAttribute("first_column", String.valueOf(firstCellReference.getCol()+1));
                    namedRangeElement.addAttribute("first_row", String.valueOf(firstCellReference.getRow()+1));
                    namedRangeElement.addAttribute("last_column", String.valueOf(lastCellReference.getCol()+1));
                    namedRangeElement.addAttribute("last_row", String.valueOf(lastCellReference.getRow()+1));
                    
                    namedRangeElement.addElement("name").setText(name.getNameName());
                    namedRangeElement.addElement("sheet_name").setText(name.getSheetName());
                    
                    namedRangeElement.addElement("refers_to_formula").setText(formula);
                }
            }            
            catch(RuntimeException e) {
            	//caused by an not implemented error in POI related to macros, and some invalid formala's that dont' relate to contiguous ranges.
            }                                    
		}		
	}

	private void sheetToXML(ArrayList<LinkedList<Element>> styleMap, int sheetIndex,
			Element sheetElement, Sheet sheet) {
		sheetElement.addAttribute("name", sheet.getSheetName());
		sheetElement.addAttribute("index", String.valueOf(sheetIndex + 1));
		sheetElement.addAttribute("hidden",
				String.valueOf(poiWorkbook.isSheetHidden(sheetIndex)));
		sheetElement.addAttribute("very_hidden",
				String.valueOf(poiWorkbook.isSheetVeryHidden(sheetIndex)));
		
		Element validations = sheetElement.addElement("data_validations");
		dataValidationsToXML(validations,sheet);

		// Columns (for column widths - styling)
		int lastCol = 1;
		int firstCol = 1;
		Element columnsElement = sheetElement.addElement("columns");

		int firstRow = sheet.getFirstRowNum();
		int lastRow = sheet.getLastRowNum();

		Element rowsElement = sheetElement.addElement("rows");
		rowsElement.addAttribute("first_row", String.valueOf(firstRow + 1));
		rowsElement.addAttribute("last_row", String.valueOf(lastRow + 1));

		for (int y = firstRow; y <= lastRow; y++) {
			Row row = sheet.getRow(y);
			if (row != null) {
				Element rowElement = rowsElement.addElement("row");
				rowElement.addAttribute("index", String.valueOf(y + 1));
				// Get height of row, if different from default
				if (sheet.getDefaultRowHeightInPoints() != row
						.getHeightInPoints())
					rowElement.addAttribute("height",
							"" + row.getHeightInPoints() + "pt");

				int firstCell = row.getFirstCellNum();

				// Skip if row has no cells
				if (firstCell == -1)
					continue;

				if (firstCell < firstCol)
					firstCol = firstCell + 1;// Number of columns

				int lastCell = row.getLastCellNum();
				if (lastCell > lastCol)
					lastCol = lastCell;// Number of columns

				for (int x = firstCell; x <= lastCell; x++) {
					Cell cell = row.getCell(x);
					if (cell != null) {
						CellInfo info = new CellInfo(cell,poiWorkbook);

						if (info.value != null) {
							Element cellElement = rowElement
									.addElement("cell");
							cellElement.addAttribute("column",
									String.valueOf(x + 1));
							cellElement.addAttribute("column_alpha",
									column_alpha(x));
							cellElement.addAttribute("row",
									String.valueOf(y + 1));
							cellElement.addAttribute("type", info.type);

							// Cell style
							// Add to style linked list
							int styleIndex = cell.getCellStyle().getIndex();

							styleMap.get(styleIndex).add(cellElement);
							cellElement.addAttribute("style",
									("style" + cell.getCellStyle()
											.getIndex()));

							if (info.formula != null) {
								cellElement.addAttribute("formula",
										info.formula);
							}
							cellElement.setText(info.value);
						}
					}
				}
			}
		}

		columnsElement.addAttribute("first_column",
				String.valueOf(firstCol));
		columnsElement.addAttribute("last_column", String.valueOf(lastCol));
		for (int x = firstCol - 1; x < lastCol; x++) {
			Element columnElement = columnsElement.addElement("column");
			columnElement.addAttribute("index", String.valueOf(x + 1));
			columnElement.addAttribute("column_alpha",
					String.valueOf(column_alpha(x)));
			columnElement.addAttribute("width",
					String.valueOf(sheet.getColumnWidth(x)));
		}
	}

	private void dataValidationsToXML(Element validations, Sheet sheet) {
		if (sheet instanceof HSSFSheet) {
			hssfDataValidationsToXML(validations,(HSSFSheet)sheet);
		}
		else {
			xssfDataValidationsToXML(validations,(XSSFSheet)sheet);
		}		
	}

	private void xssfDataValidationsToXML(Element validations, XSSFSheet sheet) {
		List<XSSFDataValidation> validationData = sheet.getDataValidations();
		for (XSSFDataValidation validation : validationData) {
			for (CellRangeAddress address : validation.getRegions().getCellRangeAddresses()) {				
				String formula = validation.getValidationConstraint().getFormula1();
				if (formula!=null) {
					Element validationEl = validations.addElement("data_validation");
					validationEl.addAttribute("first_column",String.valueOf(address.getFirstColumn()+1));
					validationEl.addAttribute("last_column",String.valueOf(address.getLastColumn()+1));
					validationEl.addAttribute("first_row",String.valueOf(address.getFirstRow()+1));
					validationEl.addAttribute("last_row",String.valueOf(address.getLastRow()+1));					
					validationEl.addElement("constraint").setText(formula);
				}						
			}			
		}
	}

	private void hssfDataValidationsToXML(Element validations, HSSFSheet sheet) {
		List<HSSFDataValidation> validationData = PatchedPoi.getInstance().getValidationData(sheet, sheet.getWorkbook());
		for (HSSFDataValidation validation : validationData) {
			for (CellRangeAddress address : validation.getRegions().getCellRangeAddresses()) {
				Element validationEl = validations.addElement("data_validation");
				validationEl.addAttribute("first_column",String.valueOf(address.getFirstColumn()+1));
				validationEl.addAttribute("last_column",String.valueOf(address.getLastColumn()+1));
				validationEl.addAttribute("first_row",String.valueOf(address.getFirstRow()+1));
				validationEl.addAttribute("last_row",String.valueOf(address.getLastRow()+1));		
				validationEl.addElement("constraint").setText(validation.getConstraint().getFormula1());			
			}			
		}
	}

	private String column_alpha(int col) {
		String result = "";
		while (col > -1) {
			int letter = (col % 26);
			result = Character.toString((char) (letter + 65)) + result;
			col = (col / 26) - 1;
		}
		return result;
	}	
}
