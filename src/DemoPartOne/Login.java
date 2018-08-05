package DemoPartOne;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

public class Login {

	public static WebDriver driver;
	
	public static void main(String[] args) throws IOException
	{
			
		System.setProperty("webdriver.chrome.driver","D:\\Softwares\\Selenium\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
		String baseURL = "http://demo.guru99.com/v4/index.php";
		driver.get(baseURL);
		String expectedTitle = "Guru99 Bank Manager HomePage";
		
		File src = new File("LoginTest.xlsx");
		FileInputStream fis=new FileInputStream(src);
		XSSFWorkbook wb = new XSSFWorkbook(fis); 
		XSSFSheet sh= wb.getSheetAt(0);
		int col;
		for(int row=1;row<6;row++) {
			col = 0;
			String userName = sh.getRow(row).getCell(col).getStringCellValue();
			driver.findElement(By.xpath("/html/body/form/table/tbody/tr[1]/td[2]/input")).sendKeys(userName);
			col = col + 1;
			String password = sh.getRow(row).getCell(col).getStringCellValue();
			driver.findElement(By.xpath("/html/body/form/table/tbody/tr[2]/td[2]/input")).sendKeys(password);
			driver.findElement(By.xpath("/html/body/form/table/tbody/tr[3]/td[2]/input[1]")).click();
			
			// driver.switchTo().alert().accept();
			
			try {
					Alert alt = driver.switchTo().alert();
					alt.accept();
					// System.out.println("Test Case " + row + " Failed");
					col = col + 1;
					FileInputStream file = new FileInputStream(new File("LoginTest.xlsx"));
					XSSFWorkbook workbook = new XSSFWorkbook(file);
					XSSFSheet sheet = workbook.getSheetAt(0);
					Cell cell = null;
					cell = sheet.getRow(row).getCell(col);
					cell.setCellValue("Test Case Failed");
					file.close();
					
					FileOutputStream outFile =new FileOutputStream(new File("LoginTest.xlsx"));
					workbook.write(outFile);
		            outFile.close();
			        
			    }
			catch (NoAlertPresentException Ex)
			{ 
			   	String actualTitle = driver.getTitle();
				if (actualTitle.contentEquals(expectedTitle))
				{
					// System.out.println("Test Case " + row + " Passed");
			        driver.findElement(By.linkText("Log out")).click();
			        driver.switchTo().alert().accept();
			        col = col + 1;
			        FileInputStream file = new FileInputStream(new File("LoginTest.xlsx"));
					XSSFWorkbook workbook = new XSSFWorkbook(file);
					XSSFSheet sheet = workbook.getSheetAt(0);
					Cell cell = null;
					cell = sheet.getRow(row).getCell(col);
					cell.setCellValue("Test Case Passed");
					file.close();
					FileOutputStream outFile =new FileOutputStream(new File("LoginTest.xlsx"));
					workbook.write(outFile);
		            outFile.close();
			    }
			}
		}
		driver.close();
	}
}
