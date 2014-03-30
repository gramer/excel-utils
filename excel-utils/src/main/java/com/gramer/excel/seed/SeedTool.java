package com.gramer.excel.seed;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

import lombok.AccessLevel;
import lombok.NoArgsConstructor;

import com.google.common.collect.Lists;

@NoArgsConstructor(access=AccessLevel.PRIVATE)
public class SeedTool {

	private static final String EXCEL_FILE_PATH = "./src/test/resources/com/gramer/excel/seed.xlsx";

	public static void main(String[] args) throws FileNotFoundException {
		DmlStatementBuilder builder = new InsertStatementBuilder();
		builder.setTargetSheets(Lists.newArrayList("Address"));
		List<String> result = builder.build(new FileInputStream(EXCEL_FILE_PATH));

		for (String sql : result) {
			System.out.println(sql);
		}
	}

}
