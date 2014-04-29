package org.buzheng.excel;


import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * @author zany@buzheng.org
 *
 * @param <T> 行数据对象
 */
public class ExcelBuilder<T> {
	private static Logger logger = LoggerFactory.getLogger(ExcelBuilder.class);

	
	/**
	 * 定义所有要显示的列
	 */
	private List<Column> columns;
	
	/**
	 * 要输出的数据
	 */
	private List<T> data;
	
	/**
	 * 表格标题
	 */
	private String caption;
	
	/**
	 * 文件类型
	 */
	private FileType fileType = FileType.XLS;

	public ExcelBuilder() {
	}
	
	/**
	 * 导出Excel到文件
	 * 程序会根据文件名的后缀自动覆盖掉设置好的类型
	 * 
	 * @param diskFilePath 目标路径
	 * @throws IOException 
	 */
	public void toFile(String diskFilePath) throws IOException {
		
		if (StringUtils.isBlank(diskFilePath)) {
			logger.info("diskFilePath is null or empty");
			return;
		}
		
		diskFilePath = diskFilePath.trim();		
		String extension = FilenameUtils.getExtension(diskFilePath);
		if (StringUtils.isBlank(extension)) {
			logger.info("diskFilePath has no extention");
			return;
		}
		
		try {
			this.setFileType(FileType.valueOf(extension.toUpperCase()));
		} catch(IllegalArgumentException e) {
			logger.info("not supported file type: {}", extension);
			return;
		}
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(diskFilePath);
			this.toOutputStream(fos);
		} finally {
			if (fos != null) {
				fos.close();
			}
		}
	}
	
	/**
	 * 导出到字节数组
	 * 
	 * @return
	 * @throws IOException
	 */
	public byte[] toByteArray() throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		this.toOutputStream(baos);
		return baos.toByteArray();
	}

	/**
	 * 导出到输出流
	 * 
	 * @param os
	 * @throws IOException
	 */
	public void toOutputStream(OutputStream os) throws IOException {
		if (this.columns == null || this.columns.isEmpty()) {
			return;
		}		
		
		Workbook wb = null;
		
		if (FileType.XLS.equals(this.fileType)) {
			wb = new HSSFWorkbook();
		} else if (FileType.XLSX.equals(this.fileType)) {
			wb = new XSSFWorkbook();	
		}
		
		Sheet sheet = wb.createSheet("Sheet");
		
		this.createCaption(wb, sheet);
		this.createHeader(wb, sheet);
		this.createBody(wb, sheet, data);
		
	    wb.write(os);
	}
	
	// 当前的行索引
	private int rowIndex = 0;
	
	// 保存 字段是否需要 sum 的配置，已字段的索引为 key
	private Map<Integer, Double> sumConfig = new HashMap<Integer, Double>();
	// 保存字段的样式， 已字段的 dataFormat 为 key
	private Map<String, CellStyle> cellStyles = new HashMap<String, CellStyle>();
	
	/**
	 * 创建表格标题
	 * @param wb
	 * @param sheet
	 */
	private void createCaption(Workbook wb, Sheet sheet) {
		if (StringUtils.isBlank(this.caption)) {
			return;
		}
		
		sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, (this.columns.size() - 1)));
		
		Row row = sheet.createRow(rowIndex++);
		Cell cell = row.createCell(0);
		cell.setCellValue(this.caption);
		
		CellStyle style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		cell.setCellStyle(style);
	}
	
	/**
	 * 创建标题行
	 * @param sheet
	 * @return 行索引号
	 */
	private void createHeader(Workbook wb, Sheet sheet) {
		
		Row row = sheet.createRow(rowIndex++);

		for (int i = this.columns.size() - 1; i >= 0; i--) {
			Column c = this.columns.get(i);
			
			Cell cell = row.createCell(i);
			cell.setCellValue(c.getTitle());
			
			if (c.isSum()) {
				sumConfig.put(i, 0.0);
			}
			
			if (c.getDataFormat() != null && c.getDataFormat().length() > 0) {
				
				if (! cellStyles.containsKey(c.getDataFormat())) {
					DataFormat format = wb.createDataFormat();//创建格式对象  
			        CellStyle style = wb.createCellStyle();
			        style.setDataFormat(format.getFormat(c.getDataFormat()));
			        
			        cellStyles.put(c.getDataFormat(), style);
				}
			}
		}
		
	}
	
	/**
	 * 创建数据表格
	 * @param sheet
	 * @param data
	 */
	@SuppressWarnings("unchecked")
	private void createBody(Workbook wb, Sheet sheet, List<T> data) {
		if (data == null || data.size() == 0) {
			return;
		}
				
		for (T t : data) {
			Row row = sheet.createRow(rowIndex++);

			for (int i = columns.size() - 1; i >= 0; i--) {
				Column column = columns.get(i);
				
				Cell cell = row.createCell(i);
				String fieldName = column.getFieldName();
				
				Object cellValue = this.readFieldValue(t, fieldName);
								
				if (column.getFieldFormatter() != null) {
					cellValue = column.getFieldFormatter().format(cellValue, t);
				} 
				
				if (cellValue == null) {
					cell.setCellType(Cell.CELL_TYPE_BLANK);
					continue;
				}
				
				if (column.getDataFormat() != null && column.getDataFormat().length() > 0) {
					cell.setCellStyle(this.cellStyles.get(column.getDataFormat()));
				}
				
				// 数字
				if (cellValue instanceof Number) {
					double d = ((Number) cellValue).doubleValue();
					cell.setCellValue(d);
					
					Double sum = this.sumConfig.get(i);
					if (sum != null) {
						sum += d;
						this.sumConfig.put(i, sum);
					}
				}
				// 日期
				else if (cellValue instanceof Date) {
					cell.setCellValue((Date) cellValue);
				}
				// 字符串
				else {
					cell.setCellValue(cellValue.toString());
				}
				
			}
		}
		
		if (this.sumConfig.size() > 0) {
			Row row = sheet.createRow(rowIndex++);
			
			for (Map.Entry<Integer, Double> entry : this.sumConfig.entrySet()) {
				Integer index = entry.getKey();
				Cell cell = row.createCell(index);
				cell.setCellValue(entry.getValue());				
				cell.setCellStyle(cellStyles.get(columns.get(index).getDataFormat()));
			}
		}
	}
	
	private Object readFieldValue(Object target, String fieldName) {
		if (target == null) {
			return null;
		}
		
		if (target instanceof Map) {
			return ((Map) target).get(fieldName);
		}
		
		try {
			return FieldUtils.readDeclaredField(target, fieldName, true);
		} catch (IllegalAccessException e) {
			e.printStackTrace();
			return null;
		}
	}

	public List<Column> getColumns() {
		return columns;
	}

	public void setColumns(List<Column> columns) {
		this.columns = columns;
	}


	public List<T> getData() {
		return data;
	}


	public void setData(List<T> data) {
		this.data = data;
	}

	public String getCaption() {
		return caption;
	}

	public void setCaption(String caption) {
		this.caption = caption;
	}

	public FileType getFileType() {
		return fileType;
	}

	public void setFileType(FileType fileType) {
		this.fileType = fileType;
	}
	
}
