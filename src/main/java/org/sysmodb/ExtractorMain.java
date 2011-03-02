package org.sysmodb;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;

public class ExtractorMain {

	private static OptionParser options;

	public ExtractorMain(String[] args) {
		processOptions(args);
		try {
			WorkbookParser workbook = new WorkbookParser(getInputStream());
			if (options.getOutputFormat().equals("xml")) {
				System.out.println(workbook.asXML());
			}
			if (options.getOutputFormat().equals("csv")) {
				System.out.println(workbook.asCSV(options.getSheet(),options.getTrim()));
			}
					
		} catch (IOException e) {
			System.err.println("IO Error reading data: " + e.getMessage());
			e.printStackTrace();
			System.exit(-1);
		}
		System.exit(0);
	}

	public static void main(String[] args) {
		new ExtractorMain(args);
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
