package org.buzheng.commons.excel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

public class ExcelBuilderTest {
	
	@BeforeClass
	public static void setUpBeforeClass() throws Exception {
		System.out.println("======================== setUpBeforeClass");
	}

	@AfterClass
	public static void tearDownAfterClass() throws Exception {
		System.out.println("======================== tearDownAfterClass");
	}

	@Before
	public void setUp() throws Exception {
		System.out.println("======================== setUp");
	}

	@After
	public void tearDown() throws Exception {
		System.out.println("======================== tearDown");
	}

	@Test
	public void testToFile() throws IOException {
		
		
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
		
	}

	private List<User> getData() {
		List<User> users = new ArrayList<User>();
		
		User user = new User();
		user.setName("zhjiun");
		user.setAge(18);
		user.setSex(0);
		user.setBalance(1000.1);
		user.setBirthDay(new Date());
		
		users.add(user);
		
		user = new User();
		user.setName("congcat");
		user.setAge(32);
		user.setSex(1);
		user.setBalance(1020.1);
		user.setBirthDay(new Date());
		
		users.add(user);
		
		
		user = new User();
		user.setName("张三");
		user.setAge(null);
		user.setSex(1);
		user.setBalance(1040.1);
		user.setBirthDay(new Date());
		
		users.add(user);
		
		user = new User();
		user.setName("里斯");
		user.setAge(89);
		user.setSex(0);
		user.setBalance(1100.1);
		user.setBirthDay(new Date());
		
		users.add(user);
		
		for (int i = 0; i < 10000; i++) {
			user = new User();
			user.setName("三天 - " + i);
			user.setAge(i + 1);
			user.setSex(0);
			user.setBalance(1100.1);
			user.setBirthDay(new Date());
			
			users.add(user);
		}
		
		return users;
	}
	
}

class User {
	
	private String name;
	
	private Integer age;
	
	private Integer sex;
	
	private Date birthDay;
	
	private Double balance;
	
	

	public Date getBirthDay() {
		return birthDay;
	}

	public void setBirthDay(Date birthDay) {
		this.birthDay = birthDay;
	}

	public Double getBalance() {
		return balance;
	}

	public void setBalance(Double balance) {
		this.balance = balance;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public Integer getSex() {
		return sex;
	}

	public void setSex(Integer sex) {
		this.sex = sex;
	}
	
}
