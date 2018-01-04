package com.tornado.zy;

import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * EXCEL导出工具类，该工具类支持简单文档导出，自定义文档导出，支持两种单元格对齐方向，默认中间对齐，支持单元格border隐藏显示。支持全手工自定义表格.
 * <br>支持行高，列宽设置，支持为普通模式表格渲染增加自定义渲染方法.{@link #render(ExlCellRender, String...)}
 * <p><b>普通模式是指根据普通列列表及List或数组数据渲染出来的工作表，例如{@link #headers(String...)},{@link #columnHeaders(String...)},{@link #contentColumns(String...)},{@link #contentData(List)}而自定义模式则是手动模式，其方法带complex头 ，参见{@link #complexHeader(List)}等</b></p>
 * <br>
 * <p>
 * 整个excel的内容格式如下：
 * <br>===fileMark 详见{@linkplain #fileMark(String)}===
 * <br>=========title 详见{@linkplain #title(String)}=====
 * <br>=====subTitle 详见{@linkplain #subtitle(String)}=====
 * <br>========heads 详见{@linkplain #headers(String...)}=====
 * <br>=c   ==================
 * <br>=o   ================== 
 * <br>=l   ========content 详见{@linkplain #contentData(List)}==========
 * <br>=u(columnHeads)  详见{@linkplain #columnHeaders(String...)} ==============
 * <br>=m   ==================
 * <br>=n   ==================
 * <br>===========end mark comments 详见{@linkplain #comments(String)}=====================
 * </p>
 * <p>本文范例参照基础防范数据导出实现 {@link SafeBaseDataAction#getStatisticsExport(SafeBaseData, HttpServletResponse)}<p>
 * @author xlsiek
 *
 */
public class ExcelExportUtils {
	private final static String[] COLUMN_NAME_ARR = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"
			,"AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ"};
	private String fileMark = "";//文件标注，第一行合并3列
	private String title = "";//标题栏，标题栏第二行，合并列数动态。根据内容宽度而定
	private String subtitle = "";//副标题。用于一些表格有盖章情况的
	private String comments;//注释行。最后一行。无格式
	private String[] headers = null;//普通模式标题数组，
	private List<ExlRow> complexHeader = null;//自定义表头，
	private String[] columnHeaders = null;//普通模式列。注意这里的普通列是指内容的第一列。
	private List<ExlRow> complexColumnHeaders = null;//自定义首列。
	private List<ExlRow> complexContent = null;//自定义内容，一旦配置了自定义内容。将忽略原有内容
	private int beginDrawRow = 0;//默认从第1行开始绘制
	private int maxColumn = 0;//最大列数索引，计算得到,通过计算complex的第一行,如果没有head 它的长度则是列头的第一行+内容的总长
	private Map<Integer,String> fillPosition = new HashMap<>();//站位map。已被使用的map位置
	private String[] contentColumns = null;//内容列
	private List<?> contentData = null;//内容
	private Map<String,Field> fieldCash = new HashMap<>();
	private int titleRowIndex = -1;//记录标题所在行。用于合并
	private Map<Integer,Integer> columnWidthMap = new HashMap<>();//存储所有的列对应的宽度。
	private Map<String,ExlCellRender<Object>> renderMap = new HashMap<>();//renader map.用来格式化
	private int contentLineHeight = 0;//普通模式数据高度
	private boolean cellNoFormat = false;//单元格无格式。特殊需求
	private int tempFontSize = 0;//临时字体。慎用。将会导致绘制表格字体全变成这样.
	private ExcelExportUtils(){
		
	}
	
	//==================暴露公共方法区=============
	/**
	 * 调用此方法初始化
	 * @return ExcelExportUtils
	 */
	public static ExcelExportUtils c(){
		return new ExcelExportUtils();
	}
	/**
	 * 设置第一行文件标注内容。无格式
	 * @param fileMark
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils fileMark(String fileMark) {
		this.fileMark = fileMark;
		return this;
	}
	/**
	 * 设置最后一行注释行内容
	 * @param comments 注释内容
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils comments(String comments) {
		this.comments = comments;
		return this;
	}
	/**
	 * 设置标题内容，无格式
	 * @param title bi
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils title(String title) {
		this.title = title;
		return this;
	}
	/**
	 * 设置表头内容，普通表头是指一个表头名组成的数组，将已无格式组装到sheet中
	 * @param headers 表头
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils headers(String... headers) {
		this.headers = headers;
		return this;
	}
	/**
	 * 增加一个render，用来格式化特定属性
	 * @param render 表头
	 * @param property 属性
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils render(ExlCellRender render,String... property) {
		for(String item : property)
			this.renderMap.put(item, render);
		return this;
	}
	/**
	 * 设置副标题，位置位于title下的一行。
	 * @param subtitle 副标题
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils subtitle(String subtitle) {
		this.subtitle = subtitle;
		return this;
	}
	/**
	 * 设置复杂表头
	 * @param complexHeader 复杂表头
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils complexHeader(List<ExlRow> complexHeader) {
		this.complexHeader = complexHeader;
		return this;
	}
	/**
	 * 设置简单列表头，列表头一般用于特殊表格。其左侧列也存在表头的情况。这种列表头只占用一列，无格式
	 * @param columnHeaders 列表头
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils columnHeaders(String... columnHeaders) {
		this.columnHeaders = columnHeaders;
		return this;
	}
	/**
	 * 设置复杂列表头
	 * @param complexColumnHeaders 复杂列表头
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils complexColumnHeaders(List<ExlRow> complexColumnHeaders) {
		this.complexColumnHeaders = complexColumnHeaders;
		return this;
	}
	/**
	 * 设置复杂内容，该方法可完全定义一个excel工作表
	 * @param complexContent 复杂内容
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils complexContent(List<ExlRow> complexContent) {
		this.complexContent = complexContent;
		return this;
	}
	
	/**
	 * 内容列，该字段用于普通excel。是一个javabean属性数组，用来萃取指定行的数值
	 * @param contentColumns 内容列
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils contentColumns(String... contentColumns) {
		this.contentColumns = contentColumns;
		return this;
	}
	/**
	 * 内容，普通模式，用于给指定列设定值
	 * @param contentData 内容
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils contentData(List<?> contentData) {
		this.contentData = contentData;
		return this;
	}
	
	/**
	 * 仅对普通模式有效，对自定义模式无效
	 * @param height 高
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils contentRowHeight(int height){
		this.contentLineHeight = height;
		return this;
	}
	
	/**去掉内容格式<br/>
	 * 仅对普通模式有效，对自定义模式无效
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils cellNoFormat(){
		this.cellNoFormat = true;
		return this;
	}
	/**
	 * 强制设置列宽，列从0开始。单位为一个字宽
	 * @param column 列
	 * @param width 宽
	 * @return ExcelExportUtils
	 */
	public ExcelExportUtils forceColumnWidth(int column,int width){
		columnWidthMap.put(column, width);
		return this;
	}
	
	/**
	 * 修改字体大小。强制
	 * @param size
	 * @return
	 */
	@Deprecated
	public ExcelExportUtils changeContentFontSize(int size){
		this.tempFontSize = size;
		return this;
	}
	

	
	/**
	 * 输出
	 * @throws IOException 
	 */
	public void export(OutputStream out) throws IOException{
		Workbook wb   = new HSSFWorkbook();	
		Map<String, CellStyle> styles = createStyles(wb);
		Sheet sheet = wb.createSheet();
		sheet.setFitToPage(true);
		sheet.setHorizontallyCenter(true);
		
		//filemark
		createFileMark(sheet);
		
		
		//title
		createTitle(sheet, styles);
		
		//subtitle
		createSubTitle(sheet, styles);
		
		//head
		createHeader(sheet, styles);
		
		//column head列头。部分特殊需求
		createColumnHeader(sheet, styles);
		
		//content
		createContent(sheet, styles);
		
		//comments
		createComments(sheet, styles);
		
		if(titleRowIndex != -1 && maxColumn > 0){
			//合并标题行
			mergedRegionByPosition(sheet, titleRowIndex, 0, 0, maxColumn + 1, false);
		}
		
		for(Integer column : columnWidthMap.keySet()){
			sheet.setColumnWidth(column, columnWidthMap.get(column)*256);
		}
		
		sheet.getPrintSetup().setPaperSize(HSSFPrintSetup.A4_PAPERSIZE);
		sheet.setMargin(HSSFSheet.BottomMargin,( double ) 0.5 );// 页边距（下）  
		sheet.setMargin(HSSFSheet.LeftMargin,( double ) 0.1 );// 页边距（左）  
		sheet.setMargin(HSSFSheet.RightMargin,( double ) 0.1 );// 页边距（右）  
		sheet.setMargin(HSSFSheet.TopMargin,( double ) 0.5 );// 页边距（上）  
		wb.write(out);
	}
	
	//=====================================================私有方法区
	
	private void addPositionBySpan(int rowspan,int colspan,int row,int cols){
		//从行开始判断
		boolean inColsExc = false;
		for(int i = 0;i < rowspan;i++){
			int inRow = row + i;
			if(colspan > 0){
				for(int j = 0;j < colspan;j++){
					int inCols = cols + j;
					addPosition2Map(inRow, inCols  );
				}
				inColsExc = true;
			}else{
				addPosition2Map(inRow, cols);
			}
		}
		if(inColsExc){
			return;
		}
		//如果行没变化
		for(int j = 0;j < colspan;j++){
			int inCols = cols + j;
			addPosition2Map(row, inCols );
		}
		
		//如果没有合并情况
		if(rowspan == 0 && colspan == 0){
			addPosition2Map(row, cols);
		}
	}
	
	private void createFileMark(Sheet sheet ){
		// 判断是否具有文件标注.先做出来，后期扩展为Row
		if (!StringUtils.isEmpty(fileMark)) {
			Row row = sheet.createRow(beginDrawRow);
			Cell cell = row.createCell(0);
			cell.setCellValue(fileMark);
			mergedRegionByPosition(sheet, beginDrawRow, 0, 0, 3, false);
			beginDrawRow++;// 走一行
		}
	}
	private void createTitle(Sheet sheet ,Map<String, CellStyle> styles){
		if(!StringUtils.isEmpty(title)){
//			int max = maxColumn  > 0 ? maxColumn : (headers.length);
			Row row = sheet.createRow(beginDrawRow);
			row.setHeightInPoints(45);
			Cell cell = row.createCell(0);
			cell.setCellValue(title);
			cell.setCellStyle(styles.get("title"));
//			mergedRegionByPosition(sheet, beginDrawRow, 0, 0, max, false);
			if(ArrayUtils.isNotEmpty(headers) && CollectionUtils.isEmpty(complexHeader)){
				mergedRegionByPosition(sheet, beginDrawRow, 0, 0, headers.length, false);
				beginDrawRow++;
			}else{
				titleRowIndex = beginDrawRow++;//走一行
			}
		}
	}
	
	private void createComments(Sheet sheet ,Map<String, CellStyle> styles){
		if(!StringUtils.isEmpty(comments)){
			int maxRowIndex = beginDrawRow;
			if(ArrayUtils.isNotEmpty(contentColumns) || CollectionUtils.isNotEmpty(complexContent)){
				 //走的是内容的情况。index没错的。
			}else if(ArrayUtils.isNotEmpty(columnHeaders) || CollectionUtils.isNotEmpty(complexColumnHeaders)){
				//计算excel的最大行数。此时beginDrawRow指向数据content的第一行我们加上列头或内容的高度即可
					//直接计算列头的高度
					if(CollectionUtils.isNotEmpty(complexColumnHeaders)){
						maxRowIndex += complexColumnHeaders.size();
					}else{
						maxRowIndex += columnHeaders.length;
					}
			}
			Row row = sheet.createRow(maxRowIndex);
			row.setHeightInPoints(15);
			Cell cell = row.createCell(0);
			cell.setCellValue(comments);
			cell.setCellStyle(styles.get("celllnb"));
//			mergedRegionByPosition(sheet, beginDrawRow, 0, 0, max, false);
			if(ArrayUtils.isNotEmpty(headers) && CollectionUtils.isEmpty(complexHeader)){
				mergedRegionByPosition(sheet, maxRowIndex, 0, 0, headers.length, false);
			}else if(maxColumn > 0){
				//直接使用maxColumn变量绘制
				mergedRegionByPosition(sheet, maxRowIndex, 0, 0, maxColumn + 1, false);
			}
			
//			beginDrawRow++;
		}
	}
	
	private void createSubTitle(Sheet sheet ,Map<String, CellStyle> styles){
		if(!StringUtils.isEmpty(subtitle)){
			Row row = sheet.createRow(beginDrawRow);
			Cell cell = row.createCell(0);
			cell.setCellValue(subtitle);
			mergedRegionByPosition(sheet, beginDrawRow, 0, 0, 3, false);
			beginDrawRow++;// 走一行
		}
	}
	
	private void addPosition2Map(int row,int cols){
		fillPosition.put(row, (fillPosition.get(row) == null ? "," : fillPosition.get(row)) +  (cols + ","));
	}
	
	/**
	 * 绘制复杂列公共方法。
	 * @param sheet
	 * @param rows
	 * @param styles
	 * @param updateRow
	 */
	private void drawComplexColumn(Sheet sheet,List<ExlRow> rows,Map<String, CellStyle> styles,boolean updateRow){
		int rowIndex = beginDrawRow ;
		for(ExlRow exRow : rows){
			//进来了肯定是要创建一行的
			Row row = getRow(sheet,rowIndex);
			
			if(exRow.getHeight() > 0){
				row.setHeightInPoints(exRow.getHeight());
			}
			
			for(int i = 0;i < exRow.getCell().size();i++)
			{	
				int columnIndex = findPosition(rowIndex);
				ExlCell exlCell = exRow.getCell().get(i);
				Cell cell = row.createCell(columnIndex);
				//追加是否包含下划线文本
				if(exlCell.getUnderLineString() == null){
					cell.setCellValue(exlCell.getValue().toString());
				}else{
					cell.setCellValue(returnUnderLineText(sheet.getWorkbook(), exlCell.getValue().toString(), exlCell.getUnderLineString()));
				}
				
				if(exlCell.getRowspan() > 0 ||  exlCell.getColspan() > 0){
					//如果有合并行列的存在，进来
					mergedRegionByPosition(sheet, rowIndex, columnIndex, exlCell.getRowspan(), exlCell.getColspan(), exlCell.isBorder());
				}
				
				String styleFlag = "cell";
				styleFlag += exlCell.isAlignCenter() ? "c" : "l";
				styleFlag += exlCell.isBorder() ? "b" : "nb";
				cell.setCellStyle(styles.get(styleFlag));
				addPositionBySpan(exlCell.getRowspan(), exlCell.getColspan(), rowIndex, columnIndex);
				//如果存在列合并，colindex往后推
				/*if(exlCell.getColspan() > 0){
					columnIndex += exlCell.getColspan();
				}else{
					columnIndex++;
				}*/
				
				//存储列宽,仅当不是合并列的情况
				if(exlCell.getWidth() > 0){
					int colspan = exlCell.getColspan() == 0 ? 1 : exlCell.getColspan();
					int avgWidth = exlCell.getWidth()/colspan;
					for(int z = 0;z < colspan ;z++){
						columnWidthMap.put(cell.getColumnIndex() + z, avgWidth);
					}
				}
			}
			
			rowIndex++;
			
		}
		
		if(updateRow){
			beginDrawRow = rowIndex;
		}
	}
	
	private void createHeader(Sheet sheet ,Map<String, CellStyle> styles){
		if(ArrayUtils.isNotEmpty(headers) || CollectionUtils.isNotEmpty(complexHeader)){
			if(CollectionUtils.isNotEmpty(complexHeader)){
				//复杂表头开始..
				drawComplexColumn(sheet, complexHeader, styles, true);
				
				addMaxColumnVar(true, beginDrawRow - complexHeader.size());
				//复杂head
			}else if(ArrayUtils.isNotEmpty(headers)){
				//简单head,只支持单行
				Row row = sheet.createRow(beginDrawRow);
				for(int i = 0;i < headers.length;i++){
					Cell cell = row.createCell(i);
					cell.setCellValue(headers[i]);
					cell.setCellStyle(styles.get("cellcb"));
				}
				beginDrawRow++;
			}
		}
	}
	
	/**
	 * 获取指定索引的row，如果不存在则新建,注意调用场合
	 * @param rowIndex 行索引
	 * @return row
	 */
	private Row getRow(Sheet sheet,int rowIndex){
		Row row = fillPosition.containsKey(rowIndex) ? sheet.getRow(rowIndex)
				: sheet.createRow(rowIndex);
		return row == null ? sheet.createRow(rowIndex) : row;
	}
	
	private void createColumnHeader(Sheet sheet, Map<String, CellStyle> styles) {
		if (ArrayUtils.isNotEmpty(columnHeaders) || CollectionUtils.isNotEmpty(complexColumnHeaders)) {
			if (CollectionUtils.isNotEmpty(complexColumnHeaders)) {
				
				drawComplexColumn(sheet, complexColumnHeaders, styles, false);
				
				//如果没有header。我们需要手动计算column数
				addMaxColumnVar(false, beginDrawRow);
				
				// 复杂head
			} else if (ArrayUtils.isNotEmpty(columnHeaders)) {
				// 简单head,只支持单列

				for (int i = 0; i < columnHeaders.length; i++) {
					Row row = sheet.createRow(beginDrawRow + i);
					Cell cell = row.createCell(0);
					cell.setCellValue(columnHeaders[i]);
					cell.setCellStyle(styles.get("cellcb"));
					addPosition2Map(beginDrawRow + i, 0);// 标识某行的第0列被占用
				}
				// 因为是生成列头。我们不需要移动行指针。行指针依然定格在表头下一行
				
				//如果没有header。我们需要手动计算column数
				if(ArrayUtils.isEmpty(headers) && CollectionUtils.isEmpty(complexHeader)){
					maxColumn += 1;
				}
			}
		}
	}
	
	private void addMaxColumnVar(boolean inCreateHeader,int nowRowIndex){
		if((ArrayUtils.isEmpty(headers) && CollectionUtils.isEmpty(complexHeader)) || inCreateHeader){
			maxColumn = Integer.parseInt(fillPosition.get(nowRowIndex).replaceFirst(".*,(\\d+),$", "$1"));
		}
	}

	private void createContent(Sheet sheet, Map<String, CellStyle> styles) {
		if (ArrayUtils.isNotEmpty(contentColumns) || CollectionUtils.isNotEmpty(complexContent)) {

			if (CollectionUtils.isNotEmpty(complexContent)) {
				drawComplexColumn(sheet, complexContent, styles, true);
				
				addMaxColumnVar(false, beginDrawRow - complexContent.size());
				// 复杂内容
			} else if (ArrayUtils.isNotEmpty(contentColumns) && CollectionUtils.isNotEmpty(contentData)) {
				// 简单head,只支持单列
				// 根据List的行数，来生成行
				int row_ = 0;
				for (Object line : contentData) {
					Row row = getRow(sheet,beginDrawRow);
					if(row.getHeightInPoints() == 12.75 && contentLineHeight > 0){
						row.setHeightInPoints(contentLineHeight);
					}
					int columnIndex = findPosition(beginDrawRow);// 找寻可使用的列位置
					for (int i = 0; i < contentColumns.length; i++) {
						Cell cell = row.createCell(columnIndex++);
						cell.setCellValue(getTarBeanValueByProperty(row_,line, contentColumns[i]));
						if(!cellNoFormat) cell.setCellStyle(styles.get("cellcb"));
					}

					beginDrawRow++;
					row_++;
				}
				
				//如果没有header。我们需要手动计算column数
				if(ArrayUtils.isEmpty(headers) && CollectionUtils.isEmpty(complexHeader)){
					maxColumn += contentColumns.length - 1;
				}

			}

		}
	}
	
	private int findPosition(int row){
		int position = -1;
		String key = null;
		do{
			key = "," + ++position + ",";
		}while(fillPosition.containsKey(row) && fillPosition.get(row).contains(key));
		return position;
	}
	
	private String getTarBeanValueByProperty(int rowIndex,Object obj,String property){
		try 
		{	
			String result = null;
			if(obj instanceof Map){
				result = String.valueOf(((Map) obj).get(property));
			}else{
				Field field = fieldCash.get(property);
				if(field == null){
					try{
						field = obj.getClass().getDeclaredField(property);
					}catch(NoSuchFieldException e){
						e.printStackTrace();
						field = obj.getClass().getSuperclass().getDeclaredField(property);
					}
					fieldCash.put(property, field);
				}
				field.setAccessible(true);
				result =  String.valueOf(field.get(obj));
			}
			if(result == null || "null".equals(result)) result = "";//去掉空值
			return renderMap.containsKey(property) ? renderMap.get(property).format(result,obj,rowIndex) : result;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
	
	private void mergedRegionByPosition(Sheet sheet,int beginRow,int beginColumn,int rowspan,int colspan,boolean hasBorder){
		if(rowspan == 0) rowspan = 1;
		if(colspan == 0) colspan = 1;
		//1代表自身.
		String exp = "$" + (COLUMN_NAME_ARR[beginColumn]) + "$" + (beginRow + 1) + ":$" + (COLUMN_NAME_ARR[beginColumn + colspan - 1]) + "$" + (beginRow + rowspan);
		CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf(exp);
		sheet.addMergedRegion(cellRangeAddress);
		if(hasBorder){
			RegionUtil.setBorderTop(1, cellRangeAddress, sheet);
	        RegionUtil.setBorderLeft(1, cellRangeAddress, sheet);
	        RegionUtil.setBorderBottom(1, cellRangeAddress, sheet);
	        RegionUtil.setBorderRight(1, cellRangeAddress, sheet);
		}
	}
	
	private HSSFRichTextString returnUnderLineText(Workbook wb,String value,String[] uStrs){
		Font font = wb.createFont();
		font.setUnderline(Font.U_SINGLE); // 下划线
		if (tempFontSize > 0) {
			font.setFontHeightInPoints((short) tempFontSize);
		}
		HSSFRichTextString richString = null;
		richString = new HSSFRichTextString(value);
		int preIndex = -1;
		Font tempFont = wb.createFont();
		if (tempFontSize > 0) {
			tempFont.setFontHeightInPoints((short)tempFontSize);
		}
		richString.applyFont(tempFont);
		for(String str : uStrs){
			int b = value.indexOf(str,preIndex);
			int e = str.length();
			preIndex = b + e;
			richString.applyFont(b, b + e, font);
		}
		
		return richString;
	}
	private Map<String, CellStyle> createStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		CellStyle style;
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		Font titleFont = wb.createFont();
		titleFont.setFontHeightInPoints((short) 18);
		style.setFont(titleFont);
		styles.put("title", style);
		
		Font tempFont = null;
		if(tempFontSize > 0){
			tempFont = wb.createFont();
			tempFont.setFontHeightInPoints((short) tempFontSize);
		}
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setWrapText(true);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put("cellcb", style);
		

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setWrapText(true);
		styles.put("cellcnb", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setWrapText(true);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put("celllb", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setWrapText(true);
		styles.put("celllnb", style);

		if(tempFont != null){
			styles.get("cellcb").setFont(tempFont);
			styles.get("cellcnb").setFont(tempFont);
			styles.get("celllb").setFont(tempFont);
			styles.get("celllnb").setFont(tempFont);
		}
		
		return styles;
	}
		
		public static void main(String[] args) { 
		ExcelExportUtils utils = ExcelExportUtils.c()
				.fileMark("附件")
				.title(  "中小学幼儿园安全防范基础数据统计表")
				.subtitle("填报单位（盖章）：");
//构建自定义表头
List<ExlRow> headers = new ArrayList<>(); 
ExlRow headRow1 = ExlRow.c().height(20);
headRow1.addCell(ExlCell.c("项目", 2, 3));
headRow1.addCell(ExlCell.c("总数", 2, 0));
headRow1.addCell(ExlCell.c("学校性质", 0, 2));
headRow1.addCell(ExlCell.c("学校类型", 0, 8));
headRow1.addCell(ExlCell.c("所在区域", 0, 3));
headers.add(headRow1);

ExlRow headRow2 = ExlRow.c().height(70);
headRow2.addCell(ExlCell.c("公办", 1, 1).width(10));
headRow2.addCell(ExlCell.c("民办", 1, 0));
headRow2.addCell(ExlCell.c("普通中专", 0, 1));
headRow2.addCell(ExlCell.c("高中(含完中)"));
headRow2.addCell(ExlCell.c("初中(含九年一贯制)"));
headRow2.addCell(ExlCell.c("职业中学"));
headRow2.addCell(ExlCell.c("小学"));
headRow2.addCell(ExlCell.c("特殊教育学校"));
headRow2.addCell(ExlCell.c("幼儿园"));
headRow2.addCell(ExlCell.c("教学点"));
headRow2.addCell(ExlCell.c("城区(含县政府所在地)"));
headRow2.addCell(ExlCell.c("乡镇"));
headRow2.addCell(ExlCell.c("农村"));
headers.add(headRow2);

//utils.complexHeader(headers);

//构建自定义列表头
List<ExlRow> compleColumnHeaders = new ArrayList<>();
ExlRow columnHeadRow1 = ExlRow.c().height(20);
columnHeadRow1.addCell(ExlCell.c("基本情况", 5	,0).width(3));
//columnHeadRow1.addCell(ExlCell.c("学校数（所）	", 0	,2).width(30));

ExlRow columnHeadRow2 = ExlRow
.c().height(20)
.addCell(ExlCell.c("在校学生数(人)", 0	,2));

ExlRow columnHeadRow3 = ExlRow
.c().height(20)
.addCell(ExlCell.c("专任教师数(人)", 0	,2));

ExlRow columnHeadRow4 = ExlRow.c().height(20);
columnHeadRow4.addCell(ExlCell.c("占地面积(㎡)", 0	,2));
ExlRow columnHeadRow5 = ExlRow.c().height(20);
columnHeadRow5.addCell(ExlCell.c("校舍建筑面积(㎡)", 0	,2));

ExlRow columnHeadRow6 = ExlRow.c().height(20);
columnHeadRow6.addCell(ExlCell.c("人防", 5	,0));
columnHeadRow6.addCell(ExlCell.c("保卫(综治)组织", 2	,0));
columnHeadRow6.addCell(ExlCell.c("有(所)", 0	,0));

ExlRow columnHeadRow6_ = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)", 0	,0));

ExlRow columnHeadRow7 = ExlRow
.c().height(20)
.addCell(ExlCell.c("专业保安数", 0	,2));

ExlRow columnHeadRow8 = ExlRow.c().height(20);
columnHeadRow8.addCell(ExlCell.c("进屋机构", 2	,0));
columnHeadRow8.addCell(ExlCell.c("有(所)", 0	,0));

ExlRow columnHeadRow9 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)", 0	,0));

ExlRow columnHeadRow10 = ExlRow.c().height(20);
columnHeadRow10.addCell(ExlCell.c("物防",12,0));
columnHeadRow10.addCell(ExlCell.c("围墙",2,0));
columnHeadRow10.addCell(ExlCell.c("有(所)"));

ExlRow columnHeadRow11 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)"));

ExlRow columnHeadRow12 = ExlRow.c().height(20);
columnHeadRow12.addCell(ExlCell.c("校门",2,0));
columnHeadRow12.addCell(ExlCell.c("有(所)"));

ExlRow columnHeadRow13 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)"));

ExlRow columnHeadRow14 = ExlRow.c().height(20);
columnHeadRow14.addCell(ExlCell.c("防卫器具数（个）",5,0));
columnHeadRow14.addCell(ExlCell.c("防割手套"));

ExlRow columnHeadRow15 = ExlRow
.c().height(20)
.addCell(ExlCell.c("钢叉"));

ExlRow columnHeadRow16 = ExlRow
				.c()
				.addCell(ExlCell.c("橡胶伸缩棍"));
ExlRow columnHeadRow17 = ExlRow
.c().height(20)
.addCell(ExlCell.c("头盔"));

ExlRow columnHeadRow18 = ExlRow
.c().height(20)
.addCell(ExlCell.c("盾牌"));


ExlRow columnHeadRow19 = ExlRow.c().height(20);
columnHeadRow19.addCell(ExlCell.c("消防设施数（个）",3,0));
columnHeadRow19.addCell(ExlCell.c("消防栓"));

ExlRow columnHeadRow20 = ExlRow
.c().height(20)
.addCell(ExlCell.c("干粉灭火器"));

ExlRow columnHeadRow21 = ExlRow
.c().height(20)
.addCell(ExlCell.c("应急灯"));

ExlRow columnHeadRow22 = ExlRow.c().height(20);
columnHeadRow22.addCell(ExlCell.c("技防",7,0));
columnHeadRow22.addCell(ExlCell.c("视频监控探头数（个）",0,2));

ExlRow columnHeadRow23 = ExlRow.c().height(20);
columnHeadRow23.addCell(ExlCell.c("红外报警装置",2,0));
columnHeadRow23.addCell(ExlCell.c("有(所)"));

ExlRow columnHeadRow24 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)"));

ExlRow columnHeadRow25 = ExlRow.c().height(20);
columnHeadRow25.addCell(ExlCell.c("警灯",2,0));
columnHeadRow25.addCell(ExlCell.c("有(所)"));

ExlRow columnHeadRow26 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)"));

ExlRow columnHeadRow27 = ExlRow.c().height(20);
columnHeadRow27.addCell(ExlCell.c("报警电话",2,0));
columnHeadRow27.addCell(ExlCell.c("有(所)"));

ExlRow columnHeadRow28 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)"));

ExlRow columnHeadRow29 = ExlRow
.c().height(20)
.addCell(ExlCell.c("无(所)",0,20));

compleColumnHeaders.add(columnHeadRow1);
compleColumnHeaders.add(ExlRow.c());
compleColumnHeaders.add(ExlRow.c());
compleColumnHeaders.add(ExlRow.c());
compleColumnHeaders.add(ExlRow.c());
compleColumnHeaders.add(columnHeadRow2);
//compleColumnHeaders.add(columnHeadRow3);
//compleColumnHeaders.add(columnHeadRow4);
//compleColumnHeaders.add(columnHeadRow5);
//compleColumnHeaders.add(columnHeadRow6);
//compleColumnHeaders.add(columnHeadRow6_);
//compleColumnHeaders.add(columnHeadRow7);
//compleColumnHeaders.add(columnHeadRow8);
//compleColumnHeaders.add(columnHeadRow9);
//compleColumnHeaders.add(columnHeadRow10);
//compleColumnHeaders.add(columnHeadRow11);
//compleColumnHeaders.add(columnHeadRow12);
//compleColumnHeaders.add(columnHeadRow13);
//compleColumnHeaders.add(columnHeadRow14);
//compleColumnHeaders.add(columnHeadRow15);
//compleColumnHeaders.add(columnHeadRow16);
//compleColumnHeaders.add(columnHeadRow17);
//compleColumnHeaders.add(columnHeadRow18);
//compleColumnHeaders.add(columnHeadRow19);
//compleColumnHeaders.add(columnHeadRow20);
//compleColumnHeaders.add(columnHeadRow21);
//compleColumnHeaders.add(columnHeadRow22);
//compleColumnHeaders.add(columnHeadRow23);
//compleColumnHeaders.add(columnHeadRow24);
//compleColumnHeaders.add(columnHeadRow25);
//compleColumnHeaders.add(columnHeadRow26);
//compleColumnHeaders.add(columnHeadRow27);
//compleColumnHeaders.add(columnHeadRow28);
//compleColumnHeaders.add(columnHeadRow29);

utils.complexColumnHeaders(compleColumnHeaders);
//utils.comments("拉斯大陆开发建设东路附近阿斯利康打飞机as姐啊送来的房价阿隆索的肌肤");
try{
	utils.export(new FileOutputStream("e:/dldl.xls"));
} catch (FileNotFoundException e) {
	e.printStackTrace();
} catch (IOException e) {
	// TODO Auto-generated catch block
	e.printStackTrace();
}
		}
		
}
