package com.gramer.excel;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;

@Controller
public class ExcelExportController {
	
	enum Browser {
		MSIE, Chrome, Opera, Firefox, Etc
	}
	
	@RequestMapping(value = "/excel", method = RequestMethod.POST)
	public void exportExcel(@RequestParam("htmlData") String htmlData, 
							@RequestParam(value = "fileName", defaultValue = "export") String fileName, 
							HttpServletRequest req, 
							HttpServletResponse res) throws IOException {
		HtmlToExcelExporter exporter = new HtmlToExcelExporter();
		res.setHeader("Content-Disposition", "attachment; filename=" + getDisposition(fileName, req) + ".xlsx");
		res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		exporter.export(URLDecoder.decode(htmlData, "UTF-8"), res.getOutputStream());
		res.flushBuffer();
	}

	private String getDisposition(String filename, HttpServletRequest req) throws UnsupportedEncodingException {
		Browser type = getBrowser(req);

		switch (type) {
		case MSIE:
			return URLEncoder.encode(filename, "UTF-8").replaceAll("\\+", "%20");
		case Chrome:
			return "\"" + new String(filename.getBytes("UTF-8"), "8859_1") + "\"";
		case Opera:
			return "\"" + new String(filename.getBytes("UTF-8"), "8859_1") + "\"";
		case Firefox:
			StringBuffer sb = new StringBuffer();
			for (int i = 0; i < filename.length(); i++) {
				char c = filename.charAt(i);
				if (c > '~') {
					sb.append(URLEncoder.encode("" + c, "UTF-8"));
				} else {
					sb.append(c);
				}
			}
			return sb.toString();

		default:
			return filename;
		}
	}
	
	private Browser getBrowser(HttpServletRequest req) {
		String header = req.getHeader("User-Agent");
		for (Browser type : Browser.values()) {
			if (header.contains(type.toString())) {
				return type;
			}
		}
		
		return Browser.Etc;
	}

}
