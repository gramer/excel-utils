package com.gramer.excel.seed;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Setter @Getter
@ToString
@AllArgsConstructor
public class ColumnMeta {
	
	private String name;
	private DataType type;
	
}
