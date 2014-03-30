package com.gramer.excel.seed;

import static org.junit.Assert.assertThat;
import static org.junit.matchers.JUnitMatchers.containsString;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

import org.junit.Test;

import com.google.common.collect.Lists;
public class SeedToolTest {

	private static final String EXCEL_FILE_PATH = "./src/test/resources/com/gramer/excel/seed.xlsx";

	@Test
	public void createInsertSql() throws FileNotFoundException {
		DmlStatementBuilder builder = new InsertStatementBuilder();
		builder.setTargetSheets(Lists.newArrayList("Address"));
		List<String> result = builder.build(new FileInputStream(EXCEL_FILE_PATH));
		
		String insertSql = result.get(0);
		assertThat(insertSql, containsString("insert into Address"));
	}

}
