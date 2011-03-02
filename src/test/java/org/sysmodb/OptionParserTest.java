package org.sysmodb;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

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
	public void testFilename() throws Exception {
		String[] args = new String[] {};
		OptionParser p = new OptionParser(args);
	}

	@Test(expected = InvalidOptionException.class)
	public void testBadFormat() throws Exception {
		String[] args = new String[] { "-o", "pdf" };
		OptionParser p = new OptionParser(args);
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
