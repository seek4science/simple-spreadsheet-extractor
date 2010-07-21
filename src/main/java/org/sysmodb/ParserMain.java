package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public class ParserMain {

	private static OptionParser options;

	public ParserMain(String[] args) {
		processOptions(args);
		try {
			WorkbookParser workbook = new WorkbookParser(getInputStream());
		} catch (IOException e) {
			System.err.println("IO Error reading data: " + e.getMessage());
			e.printStackTrace();
			System.exit(-1);
		}
	}

	public static void main(String[] args) {
		new ParserMain(args);
	}

	private InputStream getInputStream() throws IOException {
		if (options.getFilename() != null) {
			URL url = new URL("file://" + options.getFilename());
			return url.openStream();
		} else { // get from stdin
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
