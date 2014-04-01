excel
=====

一个基于POI的生成Excel的工具类。只需要对需要到处的Excel的各个列进行代码端的配置，实体对象目前只支持java bean，还不支持Map。

使用方法参考 测试类 src/test/java。大体代码段如下：

```
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
