package org.sysmodb.ParserMain;

import org.junit.Test;

import static org.junit.Assert.assertEquals;;

public class OptionParserTest {
	
	@Test
	public void testFormat() throws Exception {
		String[] args=new String[]{"-o","xml"};
		OptionParser p = new OptionParser(args);
		assertEquals("xml",p.getOutputFormat());
	}
	
	@Test(expected=InvalidOptionException.class)
	public void testBadFormat() throws Exception {
		String[] args=new String[]{"-o","pdf"};
		OptionParser p = new OptionParser(args);		
	}
	
	@Test
	public void testDefaultFormat() throws Exception {
		String[] args=new String[]{};
		OptionParser p = new OptionParser(args);
		assertEquals("xml",p.getOutputFormat());
	}
}
