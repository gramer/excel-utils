import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.junit.Test;

import com.google.common.base.Charsets;
import com.google.common.io.Files;
import com.gramer.excel.HtmlToExcelExporter;


public class HtmlExcelUtilsTest {

	@Test
	public void test() throws IOException {
		HtmlToExcelExporter h = new HtmlToExcelExporter();
		String fullPath = getClass().getResource("bizwinTable.html").getPath();
		String content = Files.toString(new File(fullPath), Charsets.UTF_8);
		FileOutputStream fos = new FileOutputStream("bizwinTable.xlsx");
		h.export(content, fos);
	}

}
