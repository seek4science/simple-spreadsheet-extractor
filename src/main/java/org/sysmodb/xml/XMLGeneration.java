package org.sysmodb.xml;

import java.io.IOException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;

import javax.xml.stream.XMLOutputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xerces.util.XMLChar;
import org.apache.xerces.xni.XMLString;
import org.sysmodb.CellInfo;

public class XMLGeneration {
	
	private final static Logger logger = Logger.getLogger(XMLGeneration.class);

	private final Workbook poiWorkbook;
	private XMLStyleHelper styleHelper = null;
	private List<CellStyle> styles = new ArrayList<CellStyle>();

	public XMLGeneration(Workbook poiWorkbook) {
		this.poiWorkbook = poiWorkbook;
		if (poiWorkbook instanceof XSSFWorkbook) {
			styleHelper = new XSSFXMLStyleHelper();
		} else {
			styleHelper = new HSSFXMLStyleHelper((HSSFWorkbook) poiWorkbook);
		}
	}

	public void outputToWriter(Writer outputWriter) throws IOException,
			XMLStreamException {
		XMLOutputFactory factory = XMLOutputFactory.newInstance();
		XMLStreamWriter xmlwriter = factory.createXMLStreamWriter(outputWriter);
		xmlwriter.writeStartDocument("1.0");
		streamXML(xmlwriter);
		xmlwriter.writeEndDocument();

		xmlwriter.flush();
		xmlwriter.close();
	}

	private void streamXML(XMLStreamWriter xmlWriter) throws XMLStreamException {
		xmlWriter.writeStartElement("workbook");
		xmlWriter
				.writeDefaultNamespace("http://www.sysmo-db.org/2010/xml/spreadsheet");
		writeNamedRanged(xmlWriter);
		writeStyles(xmlWriter);
		writeSheets(xmlWriter);
		xmlWriter.writeEndElement();
	}

	private void writeSheets(XMLStreamWriter xmlWriter)
			throws XMLStreamException {

		for (short i = 0; i < poiWorkbook.getNumberOfSheets(); i++) {
			Sheet sheet = poiWorkbook.getSheetAt(i);
			writeSheet(xmlWriter, i, sheet);
		}
	}

	private void writeSheet(XMLStreamWriter xmlWriter, short sheetIndex,
			Sheet sheet) throws XMLStreamException {
		xmlWriter.writeStartElement("sheet");
		xmlWriter.writeAttribute("name", sheet.getSheetName());
		xmlWriter.writeAttribute("index", String.valueOf(sheetIndex + 1));
		xmlWriter.writeAttribute("hidden",
				String.valueOf(poiWorkbook.isSheetHidden(sheetIndex)));
		xmlWriter.writeAttribute("very_hidden",
				String.valueOf(poiWorkbook.isSheetVeryHidden(sheetIndex)));

		writeDataValidations(xmlWriter, sheet);

		writeColumns(xmlWriter, sheet);

		writeRows(xmlWriter, sheet);

		xmlWriter.writeEndElement();

	}

	private void writeDataValidations(XMLStreamWriter xmlWriter, Sheet sheet)
			throws XMLStreamException {
		xmlWriter.writeStartElement("data_validations");
		if (sheet instanceof HSSFSheet) {
			writeHSSFDataValidations(xmlWriter, (HSSFSheet) sheet);
		} else {
			writeXSSFDataValidations(xmlWriter, (XSSFSheet) sheet);
		}
		xmlWriter.writeEndElement();
	}

	private void writeHSSFDataValidations(XMLStreamWriter xmlWriter,
			HSSFSheet sheet) throws XMLStreamException {
		try {
			List<HSSFDataValidation> validationData = sheet.getDataValidations();
			for (HSSFDataValidation validation : validationData) {
				for (CellRangeAddress address : validation.getRegions()
						.getCellRangeAddresses()) {
					String formula = validation.getValidationConstraint()
							.getFormula1();
					if (formula != null) {
						writeDataValidation(xmlWriter, address, formula);
					}
				}
			}
		}
		catch(IllegalStateException e) {
			logger.warn("Problem reading data validation table " + e.getMessage());
		}
	}

	private void writeDataValidation(XMLStreamWriter xmlWriter,
			CellRangeAddress address, String formula) throws XMLStreamException {
		xmlWriter.writeStartElement("data_validation");
		xmlWriter.writeAttribute("first_column",
				String.valueOf(address.getFirstColumn() + 1));
		xmlWriter.writeAttribute("last_column",
				String.valueOf(address.getLastColumn() + 1));
		xmlWriter.writeAttribute("first_row",
				String.valueOf(address.getFirstRow() + 1));
		xmlWriter.writeAttribute("last_row",
				String.valueOf(address.getLastRow() + 1));
		xmlWriter.writeStartElement("constraint");
		xmlWriter.writeCharacters(formula);
		xmlWriter.writeEndElement();
		xmlWriter.writeEndElement();
	}

	private void writeXSSFDataValidations(XMLStreamWriter xmlWriter,
			XSSFSheet sheet) throws XMLStreamException {
		List<XSSFDataValidation> validationData = sheet.getDataValidations();
		for (XSSFDataValidation validation : validationData) {
			for (CellRangeAddress address : validation.getRegions()
					.getCellRangeAddresses()) {
				String formula = validation.getValidationConstraint()
						.getFormula1();
				if (formula != null) {
					writeDataValidation(xmlWriter, address, formula);
				}
			}
		}
	}

	private void writeColumns(XMLStreamWriter xmlWriter, Sheet sheet)
			throws XMLStreamException {
		int firstCol = 1;
		int lastCol = 1;
		// determine first and last column
		for (int y = sheet.getFirstRowNum(); y <= sheet.getLastRowNum(); y++) {
			Row row = sheet.getRow(y);
			if (row == null) {
				continue;
			}
			int firstCell = row.getFirstCellNum();
			if (firstCell == -1)
				continue;

			if (firstCell < firstCol)
				firstCol = firstCell + 1;// Number of columns

			int lastCell = row.getLastCellNum();
			if (lastCell > lastCol)
				lastCol = lastCell;// Number of columns
		}
		xmlWriter.writeStartElement("columns");
		xmlWriter.writeAttribute("first_column", String.valueOf(firstCol));
		xmlWriter.writeAttribute("last_column", String.valueOf(lastCol));
		for (int x = firstCol - 1; x < lastCol; x++) {
			xmlWriter.writeStartElement("column");
			xmlWriter.writeAttribute("index", String.valueOf(x + 1));
			xmlWriter.writeAttribute("column_alpha",
					String.valueOf(column_alpha(x)));
			xmlWriter.writeAttribute("width",
					String.valueOf(sheet.getColumnWidth(x)));
			xmlWriter.writeEndElement();
		}
		xmlWriter.writeEndElement();
	}

	private void writeRows(XMLStreamWriter xmlWriter, Sheet sheet)
			throws XMLStreamException {
		int firstRow = sheet.getFirstRowNum();
		int lastRow = sheet.getLastRowNum();
		xmlWriter.writeStartElement("rows");
		xmlWriter.writeAttribute("first_row", String.valueOf(firstRow + 1));
		xmlWriter.writeAttribute("last_row", String.valueOf(lastRow + 1));

		for (int y = firstRow; y <= lastRow; y++) {
			Row row = sheet.getRow(y);
			if (row != null) {
				writeRow(xmlWriter, y, row, sheet);
			}
		}

		xmlWriter.writeEndElement();
	}

	private void writeRow(XMLStreamWriter xmlWriter, int index, Row row,
			Sheet sheet) throws XMLStreamException {
		xmlWriter.writeStartElement("row");
		xmlWriter.writeAttribute("index", String.valueOf(index + 1));
		if (sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
			xmlWriter.writeAttribute("height", "" + row.getHeightInPoints()
					+ "pt");
		}
		if (row.getFirstCellNum() != -1) {
			writeCells(xmlWriter, row);
		}
		xmlWriter.writeEndElement();
	}

	private void writeCells(XMLStreamWriter xmlWriter, Row row)
			throws XMLStreamException {
		for (int x = row.getFirstCellNum(); x <= row.getLastCellNum(); x++) {
			Cell cell = row.getCell(x);
			if (cell != null) {
				CellInfo info = new CellInfo(cell, poiWorkbook);

				if (info.value != null) {
					xmlWriter.writeStartElement("cell");

					xmlWriter.writeAttribute("column", String.valueOf(x + 1));
					xmlWriter.writeAttribute("column_alpha", column_alpha(x));
					xmlWriter.writeAttribute("row",
							String.valueOf(row.getRowNum() + 1));
					xmlWriter.writeAttribute("type", info.type);

					int styleIndex = cell.getCellStyle().getIndex();
					if (styles.get(styleIndex) != null) {
						xmlWriter.writeAttribute("style", "style"
								+ cell.getCellStyle().getIndex());
					}

					if (info.formula != null) {
						xmlWriter.writeAttribute("formula",
								stripControlCharacters(info.formula));
					}
					xmlWriter.writeCharacters(xml10Characters(info.value));
					xmlWriter.writeEndElement();

				}
			}
		}
	}
	
	/** 
	 * 
	 * @param original
	 * @return the same String but with XML 1.0 invalid characters (like form feed) removed
	 */
	private String xml10Characters(String original) {
		//TODO: this would be better incorporating into an custom version of XMLStreamWriter.writeCharacters
		String result = "";
		for (int i=0;i<original.length();i++) {
			char c = original.charAt(i);
			if (XMLChar.isValid(c)) {
				result += c;
			}
		}
		return result;
	}

	private void writeNamedRanged(XMLStreamWriter xmlWriter)
			throws XMLStreamException {
		xmlWriter.writeStartElement("named_ranges");

		for (int i = 0; i < poiWorkbook.getNumberOfNames(); i++) {
			Name name = poiWorkbook.getNameAt(i);
			try {
				if (!name.isDeleted() && !name.isFunctionName()) {
					String formula = name.getRefersToFormula();
					AreaReference areaReference = new AreaReference(formula);
					CellReference firstCellReference = areaReference
							.getFirstCell();
					CellReference lastCellReference = areaReference
							.getLastCell();
					formula = formula.replaceAll("\\p{C}", "?");

					xmlWriter.writeStartElement("named_range");

					xmlWriter.writeAttribute("first_column",
							String.valueOf(firstCellReference.getCol() + 1));
					xmlWriter.writeAttribute("first_row",
							String.valueOf(firstCellReference.getRow() + 1));
					xmlWriter.writeAttribute("last_column",
							String.valueOf(lastCellReference.getCol() + 1));
					xmlWriter.writeAttribute("last_row",
							String.valueOf(lastCellReference.getRow() + 1));

					xmlWriter.writeStartElement("name");
					xmlWriter.writeCharacters(name.getNameName());
					xmlWriter.writeEndElement();

					xmlWriter.writeStartElement("sheet_name");
					xmlWriter.writeCharacters(name.getSheetName());
					xmlWriter.writeEndElement();

					xmlWriter.writeStartElement("refers_to_formula");
					xmlWriter.writeCharacters(stripControlCharacters(formula));
					xmlWriter.writeEndElement();

					xmlWriter.writeEndElement();
				}
			} catch (RuntimeException e) {
				// caused by an not implemented error in POI related to macros,
				// and some invalid formala's that dont' relate to contiguous
				// ranges.
			}
		}

		xmlWriter.writeEndElement();
	}

	private void writeStyles(XMLStreamWriter xmlWriter)
			throws XMLStreamException {
		xmlWriter.writeStartElement("styles");
		gatherStyles();
		for (CellStyle style : styles) {
			if (style != null) {
				XMLStyleGenerator.writeStyle(xmlWriter, style, styleHelper);
			}
		}

		xmlWriter.writeEndElement();
	}

	private void gatherStyles() {
		for (short i = 0; i < getWorkbook().getNumCellStyles(); i++) {
			try {
				CellStyle style = getWorkbook().getCellStyleAt(i);
				if (isStyleEmpty(style)) {
					styles.add(i, null);
				} else {
					styles.add(i, style);
				}
			}
			// Sometimes XSLX messes up and reports wrong number of
			// styles...
			catch (IndexOutOfBoundsException e) {
				styles.add(i, null);
				break;
			}
		}
	}

	private boolean isStyleEmpty(CellStyle style) {
		return XMLStyleGenerator.isStyleEmpty(style, styleHelper);
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

	private Workbook getWorkbook() {
		return poiWorkbook;
	}

	private String stripControlCharacters(String original) {
		return original.replaceAll("\\p{Cntrl}", "");
	}
}
