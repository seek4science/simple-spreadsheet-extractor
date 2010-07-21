package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;

public class ParserMain {

	private static OptionParser options; 
	
	public ParserMain(String[] args) {
		processOptions(args);		
		try {
			Workbook workbook = new Workbook(getInputStream());
		} catch (IOException e) {
			System.err.println("IO Error reading data: "+e.getMessage());
			System.exit(-1);
		}
	}

	public static void main(String[] args) {		
		new ParserMain(args);		
	}
	

	private InputStream getInputStream() {
		if (options.getFilename()==null) {
			//TODO: get from filename
			return null;
		}
		else { //get from stdin
			return System.in;
		}
	}

	private void processOptions(String[] args) {
		try {
			options = new OptionParser(args);
		} catch (InvalidOptionException e) {
			System.err.println("Invalid option: " + e.getMessage());
			System.exit(-1);
		}
	}

}
