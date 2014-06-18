# Excel 生成工具类

一个基于POI的生成Excel的工具类。只需要对需要导出到Excel的各个列进行代码端的配置，就可实现导出，汇总，格式化显示等功能。

## 使用方法
使用方法参考 测试类 src/test/java。大体代码段如下：

```java
List<User> users = getData();

List<Column> cs = new ArrayList<Column>();
cs.add(new Column("姓名", "name"));		
cs.add(new Column("年龄", "age"));
cs.add(new Column("性别", "sex", new FieldFormatter<Integer, User>() {
	public Object format(Integer age, User user) {
		return age == null ? "" : (age.intValue() == 0 ? "男" : "女");
	}}));

cs.add(new Column("生日", "birthDay", "yyyy-MM-dd"));

cs.add(new Column("财产", "balance", "#,##0.00", true));


ExcelBuilder<User> eb = new ExcelBuilder<User>();
eb.setData(users);
eb.setColumns(cs);
eb.setCaption("用户信息表");

eb.toFile("D:\\users.xlsx");
System.out.println("已导出");
```

## 依赖
```xml
<repositories>
	<repository>
		<id>buzheng-excel-mvn-repo</id>
		<url>https://raw.github.com/buzheng/buzheng-excel/mvn-repo/</url>
	</repository>
</repositories>

<dependency>
	<groupId>org.buzheng</groupId>
	<artifactId>buzheng-excel</artifactId>
	<version>0.1</version>
</dependency>
```
