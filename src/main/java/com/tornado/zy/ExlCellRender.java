package com.tornado.zy;

@FunctionalInterface
public interface ExlCellRender<T> {
	String format(String v, T record, int index);
}

