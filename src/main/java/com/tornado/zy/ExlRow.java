package com.tornado.zy;

import java.util.ArrayList;
import java.util.List;

/**
 * 行，代表工作表的一行记录。暂时支持高度属性
 * 
 * @author xlsiek
 *
 */
public class ExlRow {
	private int height;
	private List<ExlCell> cell = new ArrayList<>();

	private ExlRow() {

	}

	/**
	 * 初始化
	 * 
	 * @return ExlRow
	 */
	public static ExlRow c() {
		return new ExlRow();
	}

	/**
	 * 设置高度.高度为字高 pt，10个单位大概为一个字的高度
	 * 
	 * @param height 高
	 * @return ExlRow
	 */
	public ExlRow height(int height) {
		this.height = height;
		return this;
	}

	public int getHeight() {
		return height;
	}

	/**
	 * 向指定的行添加一个单元格
	 * 
	 * @param cell 单元格
	 * @return ExlRow
	 */
	public ExlRow addCell(ExlCell cell) {
		this.cell.add(cell);
		return this;
	}

	public List<ExlCell> getCell() {
		return cell;
	}
}