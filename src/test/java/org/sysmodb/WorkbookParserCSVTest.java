package org.sysmodb;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertEquals;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.StringWriter;
import java.io.Writer;
import java.net.URL;

import org.junit.Test;

public class WorkbookParserCSVTest {

	@Test
	public void testAsCSVSanity() throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class
				.getResource("/test-spreadsheet-for-csv.xls");
		assertNotNull(resourceURL);
		InputStream stream = resourceURL.openStream();
		WorkbookParser p = new WorkbookParser(stream);
		String csv = p.asCSV(1);
		assertNotNull(csv);

		String expectedResult = expectedResult("/test-spreadsheet-for-csv.csv");
		assertEquals(expectedResult,csv);
		System.out.println(expectedResult);
	}

	private String expectedResult(String resourceName) throws Exception {
		URL resourceURL = WorkbookParserCSVTest.class.getResource(resourceName);

		Writer writer = new StringWriter();

		char[] buffer = new char[1024];

		Reader reader = new BufferedReader(new InputStreamReader(
				resourceURL.openStream(), "UTF-8"));
		int n;
		while ((n = reader.read(buffer)) != -1) {
			writer.write(buffer, 0, n);
		}
		return writer.toString();
	}
}
