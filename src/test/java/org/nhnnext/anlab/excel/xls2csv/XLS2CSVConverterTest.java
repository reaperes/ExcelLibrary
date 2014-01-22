package org.nhnnext.anlab.excel.xls2csv;

import org.junit.Test;

import static org.junit.Assert.*;

public class XLS2CSVConverterTest {
	private XLS2CSVConverter converter;

	@Test
	public void testConvert() throws Exception {
		new XLS2CSVConverter().convertToCSV("src/test/resources/sample.xlsx", 0, "src/test/resources/sample.csv");
	}
	
	@Test
	public void testCharacter() throws Exception {
		String s = "";
		System.out.println((int)s.charAt(0));
		
		for (int i=0; i<128; i++) {
			char c = (char)i;
			System.out.print(i);
			System.out.print(": ");
			System.out.print(c);
			System.out.print(" ");
		}
	}
}
