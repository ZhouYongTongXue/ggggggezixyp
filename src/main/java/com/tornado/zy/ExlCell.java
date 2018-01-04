package com.tornado.zy;

/**
 * 单元格（列），多个单元格存在一个行记录中。单元格支持合并行,合并列记录。
 * 
 * @author xlsiek
 *
 */
public class ExlCell {
	private boolean border = true;
	private int rowspan = 0;
	private int colspan = 0;
	private int width;
	private boolean alignCenter = true;
	private Object value;
	
	private String[] underLineString = null;//下划线文本

	private ExlCell(Object value, int rowspan, int colspan, boolean border) {
		this.border = border;
		this.value = value;
		setRowspan(rowspan);
		setColspan(colspan);
	}

	/**
	 * 初始化
	 * 
	 * @param value
	 *            值，需要显示的内容
	 * @param rowspan
	 *            行合并数
	 * @param colspan
	 *            列合并数
	 * @param border
	 *            是否有边框
	 * @return ExlCell
	 */
	public static ExlCell c(Object value, int rowspan, int colspan, boolean border) {
		return new ExlCell(value, rowspan, colspan, border);
	}

	/**
	 * 初始化，默认有边框
	 * 
	 * @param value
	 *            值，需要显示的内容
	 * @param rowspan
	 *            行合并数
	 * @param colspan
	 *            列合并数
	 * @return ExlCell
	 */
	public static ExlCell c(Object value, int rowspan, int colspan) {
		return new ExlCell(value, rowspan, colspan, true);
	}

	/**
	 * 初始化,默认无单元格合并操作，适用于单个简单单元格添加
	 * 
	 * @param value
	 *            值，需要显示的内容
	 * @param border
	 *            是否有边框
	 * @return ExlCell
	 */
	public static ExlCell c(Object value, boolean border) {
		return new ExlCell(value, 0, 0, border);
	}

	/**
	 * 初始化,默认无单元格合并操作，有边框，适用于单个简单单元格添加
	 * 
	 * @param value
	 *            需要显示的值
	 * @return ExlCell
	 */
	public static ExlCell c(Object value) {
		return new ExlCell(value, 0, 0, true);
	}

	/**
	 * 设置左对齐
	 * 
	 * @return ExlCell
	 */
	public ExlCell alignLeft() {
		alignCenter = false;
		return this;
	}

	/**
	 * 取消边框
	 * 
	 * @return ExlCell
	 */
	public ExlCell noBorder() {
		border = false;
		return this;
	}

	public int getWidth() {
		return width;
	}

	public boolean isAlignCenter() {
		return alignCenter;
	}

	/**
	 * 设置列宽，这里注意，当列为合并列时，该宽度将平分。如果存在某列多次定义宽度以最后一次为准,2个单位为一个字的宽度
	 * 
	 * @param width 宽
	 * @return ExlCell
	 */
	public ExlCell width(int width) {
		this.width = width;
		return this;
	}
	
	/**
	 * 设置下划线，给原value部分替换成下划线
	 * @param u
	 * @return
	 */
	public ExlCell uString(String... u){
		underLineString = u;
		return this;
	}

	public boolean isBorder() {
		return border;
	}

	public int getRowspan() {
		return rowspan;
	}

	public void setRowspan(int rowspan) {
		if (rowspan <= 1)
			rowspan = 0;
		this.rowspan = rowspan;
	}

	public int getColspan() {
		return colspan;
	}

	public void setColspan(int colspan) {
		if (colspan <= 1)
			colspan = 0;
		this.colspan = colspan;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}
	
	public String[] getUnderLineString(){
		return underLineString;
	}

}