package com.gramer.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.google.common.collect.Lists;

public class HtmlToExcelExporter {

	public void export(String htmlData, OutputStream outputStream) {
		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet();
		List<CellRangeAddress> mergedRanges = Lists.newArrayList();
		int rowCount = 0;
		Row row = null;
		Document doc = Jsoup.parse(htmlData);
		for (Element table : doc.select("table")) {
			for (Element tr : table.select("tr")) {
				row = sheet.createRow(rowCount);
				buildCell(sheet, mergedRanges, rowCount, row, tr.select("th"), createHeaderStyle(wb));
				buildCell(sheet, mergedRanges, rowCount, row, tr.select("td"), null);
				rowCount++;
				sheet = wb.getSheetAt(0);
			}
		}
		
		try {
			wb.write(outputStream);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	private void buildCell(Sheet sheet, List<CellRangeAddress> mergedRanges, int rowCount, Row row, Elements elements, CellStyle cellStyle) {
		Cell cell;
		int count = 0;
		for (Element element : elements) {
			while(isMerged(mergedRanges, rowCount, count)) { count++; }
			cell = row.createCell(count);
			if (cellStyle != null)
				cell.setCellStyle(cellStyle);
			cell.setCellValue(element.text());
			if (element.hasAttr("colspan") || element.hasAttr("rowspan")) {
				int colspan = element.attr("colspan").isEmpty() ? 0 : Integer.parseInt(element.attr("colspan")) -1;
				int rowspan = element.attr("rowspan").isEmpty() ? 0 : Integer.parseInt(element.attr("rowspan")) -1;
				CellRangeAddress region = new CellRangeAddress(rowCount, rowCount + rowspan, count, count + colspan);
				mergedRanges.add(region);
				sheet.addMergedRegion(region);
				count = count + colspan;
			}
			count++;
		}
	}

	private Font createHeaderFont(Workbook wb) {
		Font headerFont = wb.createFont();
		headerFont.setBoldweight(headerFont.BOLDWEIGHT_BOLD);
		headerFont.setFontHeightInPoints((short) 12);
		return headerFont;
	}

	private CellStyle createHeaderStyle(Workbook wb) {
		CellStyle headerStyle = wb.createCellStyle();
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		headerStyle.setFont(createHeaderFont(wb));
		return headerStyle;
	}
	

	private boolean isMerged(List<CellRangeAddress> cellRangeAddresses, int rowIdx, int colIdx) {
		for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
			if (cellRangeAddress.getFirstRow() <= rowIdx && cellRangeAddress.getLastRow() >= rowIdx 
				&& (cellRangeAddress.getFirstColumn() <= colIdx && cellRangeAddress.getLastColumn() >= colIdx)) {
					return true;
				}
		}
		
		return false;
	}
}
