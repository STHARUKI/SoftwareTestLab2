package cn.edu.tju.labSelenium;


import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.*;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

public class TestBaidu {
  private WebDriver driver;
  private String baseUrl;
  private boolean acceptNextAlert = true;
  private StringBuffer verificationErrors = new StringBuffer();

  @Before
  public void setUp() throws Exception {
	  String driverPath = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chromedriver.exe";
	  System.setProperty("webdriver.chrome.driver", driverPath);
	  driver = new ChromeDriver();
	  baseUrl = "http://121.193.130.195:8800";
	  
  }

  @Test
  public void testBaidu() throws Exception {
	  String excelPath = "C:\\Users\\Administrator\\Desktop\\rj.xlsx";
	  File excel = new File(excelPath);
	  InputStream fis = new FileInputStream(excel);
	  Workbook book = null; 
	  book = new XSSFWorkbook(fis);
	  org.apache.poi.ss.usermodel.Sheet sheet1 = book.getSheetAt(0);
	  
	  File writename = new File("C:\\Users\\Administrator\\Desktop\\output.txt"); // 相对路径，如果没有则要建立一个新的output。txt文件
	  writename.createNewFile(); // 创建新文件
	  BufferedWriter out = new BufferedWriter(new FileWriter(writename));
	  
	  for(int rowNumber = 2; rowNumber <= sheet1.getLastRowNum(); ++rowNumber) {
		  Row row = sheet1.getRow(rowNumber);
		  Cell cell1 = row.getCell(1);
		  Cell cell2 = row.getCell(2);
		  Cell cell3 = row.getCell(3);
		  String studentId = cell1.toString();
		  int k = studentId.indexOf("E");
		  studentId = studentId.substring(0,1) + studentId.substring(2,k);
		  studentId = String.format("%-10s", studentId).replace(' ', '0');
		  System.out.println(studentId);
		  String studentName = cell2.getStringCellValue();
		  String studentUrl = cell3.getStringCellValue();
		  driver.get(baseUrl + "/");
		  WebElement weID = driver.findElement(By.name("id"));
		  WebElement wePassword = driver.findElement(By.name("password"));
		  weID.click();
		  driver.findElement(By.name("id")).clear();
		  driver.findElement(By.name("id")).sendKeys(studentId);
		  wePassword.click();
		  driver.findElement(By.name("password")).clear();
		  driver.findElement(By.name("password")).sendKeys(studentId.substring(4,10));
		  driver.findElement(By.id("btn_login")).click();
		  String student_id = driver.findElement(By.id("student-id")).getText();
		  String student_name = driver.findElement(By.id("student-name")).getText();
		  String student_git = driver.findElement(By.id("student-git")).getText();
		  if(studentId.equals(student_id) && student_name.equals(studentName) && student_git.equals(studentUrl)) {
			  out.write(studentId + ": OK");
			  out.newLine();
			  out.flush(); 
		  } else {
			  out.write(studentId + ": Error");
			  out.newLine();
			  out.write("Excel: StudentId:" + studentId + "StudentName:" + studentName + "StudentGit:" + studentUrl);
			  out.newLine();
			  out.write("Website: StudentId:" + student_id + "StudentName:" + student_name + "StudentGit:" + student_git);
			  out.newLine();
			  out.flush(); // 
		  }

		  driver.findElement(By.linkText("LOG OUT")).click();
	  }
	  

	  out.close(); // 最后记得关闭文件
	  book.close();

  }

  @After
  public void tearDown() throws Exception {
//    driver.quit();
//    String verificationErrorString = verificationErrors.toString();
//    if (!"".equals(verificationErrorString)) {
//      fail(verificationErrorString);
//    }
  }

  private boolean isElementPresent(By by) {
    try {
      driver.findElement(by);
      return true;
    } catch (NoSuchElementException e) {
      return false;
    }
  }

  private boolean isAlertPresent() {
    try {
      driver.switchTo().alert();
      return true;
    } catch (NoAlertPresentException e) {
      return false;
    }
  }

  private String closeAlertAndGetItsText() {
    try {
      Alert alert = driver.switchTo().alert();
      String alertText = alert.getText();
      if (acceptNextAlert) {
        alert.accept();
      } else {
        alert.dismiss();
      }
      return alertText;
    } finally {
      acceptNextAlert = true;
    }
  }
}


