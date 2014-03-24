package com.gramer.multiline;

import static org.junit.Assert.assertEquals;

import org.adrianwalker.multilinestring.Multiline;
import org.junit.Test;

public class MultilineTest {
	
	/**
	 * hello
	 */
	@Multiline
	private static final String test = null;

	@Test
	public void test() {
		assertEquals("hello", test.trim());
	}

}
