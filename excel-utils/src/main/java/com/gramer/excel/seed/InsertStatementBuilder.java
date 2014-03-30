package com.gramer.excel.seed;

import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.google.common.collect.Lists;

public class InsertStatementBuilder implements DmlStatementBuilder {

	private static final String LINE_SEPARATOR = System.getProperty("line.separator");

	private List<String> targetSheets;
	
	@Override
	public List<String> build(InputStream in) {
		List<String> result = Lists.newArrayList();
		try {
			Workbook workbook = WorkbookFactory.create(in);
			final int countOfSheets = workbook.getNumberOfSheets();
			for (int i = 0; i < countOfSheets; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				if (targetSheets.contains(sheet.getSheetName())) {
					result.addAll(build(sheet));
					result.add(LINE_SEPARATOR);
				}
			}
		} catch (Exception e) {
			throw new RuntimeException("엑셀 파일을 여는 도중 오류가 발생하였습니다.");
		}

		return result;
	}

	private List<String> build(Sheet sheet) {
		List<String> result = Lists.newArrayList();
		Row headerRow = sheet.getRow(0);
		final int lastColumnIdx = headerRow.getLastCellNum();
		final int lastRowIdx = sheet.getLastRowNum();
		List<ColumnMeta> columnMetas = createColumnMetas(headerRow, lastColumnIdx);

		for (int i = 1; i <= lastRowIdx; i++) {
			result.add(createInsertStatementString(sheet.getSheetName(), columnMetas, sheet.getRow(i), lastColumnIdx));
		}

		return result;
	}

	private List<ColumnMeta> createColumnMetas(Row headerRow, final int lastColumnIdx) {
		List<ColumnMeta> result = Lists.newArrayList();

		for (int i = 0; i < lastColumnIdx; i++) {
			String cellValue = headerRow.getCell(i).getStringCellValue();
			String[] metaStrings = cellValue.split(":");
			if (metaStrings.length == 0) {
				throw new RuntimeException(String.format("첫번째 행은 name:data_type 입니다. invalid value[%s]", cellValue));
			} else if (metaStrings.length == 1) {
				result.add(new ColumnMeta(metaStrings[0], DataType.STRING));
			} else {
				result.add(new ColumnMeta(metaStrings[0], DataType.from(metaStrings[1])));
			}
		}

		return result;
	}

	private String createInsertStatementString(String sheetName, List<ColumnMeta> metas, Row row, final int columnIndex) {
		if (row == null) {
			return "";
		}

		StringBuffer buffer = new StringBuffer();
		try {
			buffer.append("insert into ");
			buffer.append(sheetName);
			buffer.append("(");

			ColumnMeta firstColumnMeta = metas.get(0);
			buffer.append(firstColumnMeta.getName());

			for (int i = 1, limit = metas.size(); i < limit; i++) {
				buffer.append(",");
				buffer.append(metas.get(i).getName());
			}

			buffer.append(") values (");
			buffer.append(getStringByCellType(row.getCell(0)));
			for (int i = 1; i < columnIndex; i++) {
				buffer.append(",");
				buffer.append(getStringByCellType(row.getCell(i)));
			}

			buffer.append(");");
			return buffer.toString();
		} catch (Exception e) {
			System.out.println(String.format("sheetName=[%s], row=[%s]colIndex=[%d], ", sheetName, row.toString(), columnIndex));
			throw new RuntimeException(e);
		}

	}

	private String getStringByCellType(Cell cell) {
		if (cell == null) {
			return "null";
		}

		final int cellType = cell.getCellType();
		switch (cellType) {
		case Cell.CELL_TYPE_BLANK:
			return "null";
		case Cell.CELL_TYPE_STRING: {
			String value = cell.getRichStringCellValue().toString();
			return value == null || value.isEmpty() ? "null" : "'" + value + "'";
		}
		case Cell.CELL_TYPE_NUMERIC: {
			double value = cell.getNumericCellValue();
			return value == 0 ? "0" : String.valueOf(value);
		}
		default:
			throw new IllegalArgumentException(String.format("해당 셀타입[%s]은 지원하지 않습니다. ", cell.getCellType()));
		}

	}

	public void setTargetSheets(List<String> targetSheets) {
		this.targetSheets = targetSheets;
	}

}
