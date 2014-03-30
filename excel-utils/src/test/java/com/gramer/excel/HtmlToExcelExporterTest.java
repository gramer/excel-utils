package com.gramer.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.adrianwalker.multilinestring.Multiline;
import org.junit.Test;

import com.google.common.base.Charsets;
import com.google.common.io.Files;
import com.gramer.excel.html.HtmlToExcelExporter;

public class HtmlToExcelExporterTest {

	/**
	 * <table>
	 * <thead>
	 * <tr>
	 * <th rowspan='2'>accCategory</th>
	 * <th colspan='3'>2012</th>
	 * <th rowspan='2'>2012 total</th>
	 * <th colspan='3'>2013</th>
	 * <th rowspan='2'>2013 total</th>
	 * </tr>
	 * <tr>
	 * <th>1th</th>
	 * <th>2th</th>
	 * <th>3th</th>
	 * <th>1th</th>
	 * <th>2th</th>
	 * <th>3th</th>
	 * </tr>
	 * </thead> <tbody>
	 * <tr>
	 * <td rowspan='2'><div>aaa</div></td>
	 * <td>1000</td>
	 * <td>2000</td>
	 * <td>3000</td>
	 * <td>6000</td>
	 * <td>1000</td>
	 * <td>2000</td>
	 * <td>3000</td>
	 * <td>6000</td>
	 * </tr>
	 * <tr>
	 * <td>1000</td>
	 * <td>2000</td>
	 * <td>3000</td>
	 * <td>6000</td>
	 * <td>1000</td>
	 * <td>2000</td>
	 * <td>3000</td>
	 * <td>6000</td>
	 * </tr>
	 * </tbody>
	 * </table>
	 */
	@Multiline
	private static final String HTML_DATA1 = null;

	@Test
	public void craeteExcel1() throws FileNotFoundException {
		FileOutputStream fos = new FileOutputStream("test.xlsx");
		HtmlToExcelExporter exporter = new HtmlToExcelExporter();
		exporter.export(HTML_DATA1, fos);
	}

	@Test
	public void craeteExcel2() throws IOException {
		String fullPath = getClass().getResource("bizwinTable.html").getPath();
		String content = Files.toString(new File(fullPath), Charsets.UTF_8);
		FileOutputStream fos = new FileOutputStream("bizwinTable.xlsx");
		HtmlToExcelExporter exporter = new HtmlToExcelExporter();
		exporter.export(content, fos);
	}

}
