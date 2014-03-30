package com.gramer.excel.seed;

import java.io.InputStream;
import java.util.List;

public interface DmlStatementBuilder {

	List<String> build(InputStream in);

	void setTargetSheets(List<String> targetSheets);

}
