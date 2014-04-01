package org.buzheng.commons.excel;




/**
 * 字段内容数据 加工处理
 * 
 * @author Adam
 *
 * @param <R> 当前行的数据对象类型
 */
public interface FieldFormatter<F, R> {
	
	/**
	 * 对字段内容进行处理，返回自定义的内容
	 * 
	 * @param fieldValue 字段内容
	 * @param rowValue   字段所在行数据对象
	 * @return 自定义的内容
	 */
	public Object format(F fieldValue, R rowValue);
}
