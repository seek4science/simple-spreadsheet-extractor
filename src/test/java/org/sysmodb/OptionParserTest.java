package org.sysmodb;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertFalse;

import org.junit.Test;

public class OptionParserTest {

	@Test
	public void testFormat() throws Exception {
		String[] args = new String[] { "-o", "xml" };
		OptionParser p = new OptionParser(args);
		assertEquals("xml", p.getOutputFormat());

		args = new String[] { "-f", "fred.txt", "-o", "xml" };
		p = new OptionParser(args);
		assertEquals("xml", p.getOutputFormat());
		assertEquals("fred.txt",p.getFilename());
		
		args = new String[] { "-f","fred.txt","-o","csv"};
		p = new OptionParser(args);	
		assertEquals("csv",p.getOutputFormat());
	}
	
	@Test
	public void testSheet() throws Exception {
		String[] args = new String[] { "-o", "xml","-s","2" };
		OptionParser p = new OptionParser(args);
		assertEquals("xml", p.getOutputFormat());
		assertEquals(2,p.getSheet());
	}
	
	@Test(expected = InvalidOptionException.class)
	public void testBadSheet() throws Exception {
		String[] args = new String[] { "-o", "xml","-s","a word" };
		new OptionParser(args);
	}
	
	@Test
	public void testTrim() throws Exception {
		String[] args = new String[] { "-o", "csv","-t" };
		OptionParser p = new OptionParser(args);
		assertEquals("csv", p.getOutputFormat());
		assertTrue(p.getTrim());
		
		args = new String[] { "-o", "csv"};
		p = new OptionParser(args);
		assertFalse(p.getTrim());
		
		args = new String[] {"-t", "-o", "csv"};
		p = new OptionParser(args);
		assertTrue(p.getTrim());

		args = new String[] {"-o", "csv","-s","1","-t"};
		p = new OptionParser(args);
		assertTrue(p.getTrim());
	}
	
	public void testDefaultSheet() throws Exception {
		String[] args = new String[] { "-o", "xml" };
		OptionParser p = new OptionParser(args);		
		//defaults to all with XML (indicated by -1)
		assertEquals(-1,p.getSheet());
		
		args = new String[] {"-o","csv"};
		p = new OptionParser(args);	
		assertEquals(1,p.getSheet());
	}

	@Test(expected = InvalidOptionException.class)
	public void testBadFormat() throws Exception {
		String[] args = new String[] { "-o", "pdf" };
		new OptionParser(args);
	}

	@Test
	public void testDefaultFormat() throws Exception {
		String[] args = new String[] {};
		OptionParser p = new OptionParser(args);
		assertEquals("xml", p.getOutputFormat());
	}

	@Test
	public void testDefaultFilenameNull() throws Exception {
		String[] args = new String[] {};
		OptionParser p = new OptionParser(args);
		assertNull(p.getFilename());
	}
}
