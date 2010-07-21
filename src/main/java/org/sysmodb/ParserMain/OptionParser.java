package org.sysmodb.ParserMain;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class OptionParser {
	
	private String filename=null;
	private String outputFormat="xml";
	
	private static List<String> VALID_FORMATS = Arrays.asList("xml");
	
	
	public OptionParser(String [] args) throws InvalidOptionException {
		for (int i=0;i<args.length;i++) {
			String arg=args[i];
			if (arg.equals("-o")) {
				i++;
				setOutputFormat(args[i]);
				return;
			}
			
			throw new InvalidOptionException("Unrecognised option: "+args[i]);
			
		}
	}	

	public String getOutputFormat() {
		return outputFormat;
	}
	
	private void setOutputFormat(String format) throws InvalidOptionException {
		format = format.toLowerCase();
		if (VALID_FORMATS.contains(format)) {
			outputFormat=format;
		}
		else {
			throw new InvalidOptionException("Invalid output format: "+format);
		}
		
	}
	
}
