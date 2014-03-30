package com.gramer.excel.seed;

public enum DataType {
	
	NUMBER, STRING, DATE;
	
	public static DataType from(String value) {
		for (DataType each : DataType.values()) {
			if (each.name().equalsIgnoreCase(value)) {
				return each;
			}
		}
		
		throw new IllegalArgumentException(String.format("[%s] not supported value", value));
	}

}
