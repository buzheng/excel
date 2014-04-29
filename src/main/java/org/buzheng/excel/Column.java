package org.buzheng.excel;

import org.apache.commons.lang3.StringUtils;



/**
 * 表示Excel数据中的一列
 * 
 * @author zany@buzheng.org
 *
 */
public class Column {
	
	/**
	 * 列标题
	 */
	private String title;
	
	/**
	 * 列对应的 属性字段
	 */
	private String fieldName;
	
	/**
	 * 属性字段内容格式化器
	 */
	private FieldFormatter fieldFormatter;

	/**
	 * Excel 单元格格式
	 */
	private String dataFormat;
	
	/**
	 * 是否需要汇总
	 */
	private boolean sum = false;

	
	public Column(String title, String fieldName) {
		this(title, fieldName, (String) null);
	}

	public Column(String title, String fieldName, FieldFormatter formatter) {
		this(title, fieldName, formatter, null);
	}

	public Column(String title, String fieldName, boolean sum) {
		this(title, fieldName, null, null, sum);
	}

	public Column(String title, String fieldName, String dataFormat) {
		this(title, fieldName, null, dataFormat);
	}

	public Column(String title, String fieldName, FieldFormatter formatter,
			String dataFormat) {
		this(title, fieldName, formatter, dataFormat, false);
	}
		
	public Column(String title, String fieldName, String dataFormat, boolean sum) {
		this(title, fieldName, null, dataFormat, sum);
	}

	public Column(String title, String fieldName,
			FieldFormatter fieldFormatter, String dataFormat, boolean sum) {
		super();		
		this.setTitle(title);
		this.setFieldName(fieldName);
		this.setFieldFormatter(fieldFormatter);
		this.setDataFormat(dataFormat);
		this.setSum(sum);
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title == null ? "" : title;
	}

	public String getFieldName() {
		return fieldName;
	}

	public void setFieldName(String field) {
		if (StringUtils.isBlank(field)) {
			throw new IllegalArgumentException("invalid field name: required and not blank");
		}
		
		this.fieldName = field.trim();
	}

	public FieldFormatter getFieldFormatter() {
		return fieldFormatter;
	}

	public void setFieldFormatter(FieldFormatter formatter) {
		this.fieldFormatter = formatter;
	}

	public String getDataFormat() {
		return dataFormat;
	}

	public void setDataFormat(String dataFormat) {
		if (! StringUtils.isBlank(dataFormat)) {
			this.dataFormat = dataFormat;
		}
	}

	public boolean isSum() {
		return sum;
	}

	public void setSum(boolean sum) {
		this.sum = sum;
	}
	
}
