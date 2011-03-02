package org.sysmodb;

import java.util.Arrays;
import java.util.List;

public class OptionParser {

	private String filename = null;
	private String outputFormat = "xml";
	private int sheet = -1;

	private static List<String> VALID_FORMATS = Arrays.asList(new String [] {"xml","csv"});

	public OptionParser(String[] args) throws InvalidOptionException {
		for (int i = 0; i < args.length; i++) {
			String arg = args[i];
			if (arg.equals("-o")) {
				i++;
				setOutputFormat(args[i]);				
			}
			else if (arg.equals("-f")) {
				i++;
				setFilename(args[i]);				
			}
			else if (arg.equals("-s")) {
				i++;
				setSheet(args[i]);
			}
			else {
				throw new InvalidOptionException("Unrecognised option: " + args[i]);
			}
		}
		//if CSV format and sheet is not defined, then defaults to the first sheet
		if (getOutputFormat().equals("csv") && getSheet()==-1) {
			sheet=1;
		}
	}

	public int getSheet() {
		return sheet;
	}
	public String getOutputFormat() {
		return outputFormat;
	}

	private void setOutputFormat(String format) throws InvalidOptionException {
		format = format.toLowerCase();
		if (VALID_FORMATS.contains(format)) {
			outputFormat = format;
		} else {
			throw new InvalidOptionException("Invalid output format: " + format);
		}
	}

	private void setFilename(String filename) {
		this.filename = filename;
	}

	public String getFilename() {
		return filename;
	}
	
	private void setSheet(String sheet)  throws InvalidOptionException{
		try {
			this.sheet = Integer.parseInt(sheet);
		}
		catch(NumberFormatException e) {
			throw new InvalidOptionException("Invalid sheet number supplied: '"+sheet+"'");
		}
	}

}
