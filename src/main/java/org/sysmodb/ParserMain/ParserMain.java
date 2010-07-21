package org.sysmodb.ParserMain;

public class ParserMain {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		
		try {
			OptionParser options = new OptionParser(args);
		} catch (InvalidOptionException e) {
			System.err.println("Invalid option: " + e.getMessage());
			System.exit(-1);
		}
		
	}

}
