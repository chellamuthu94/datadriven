package writingexcelfile;

import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

public class writingfile {
	static public ChromeDriver driver;
@Test
	public void DF() throws InterruptedException {
		
		try {
			String fileLocation =( "C:\\Users\\FB\\eclipse-workspace\\google\\src\\main\\java\\window\\write.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Employee Data");
			int i = 2;
			int  rows=1;
			//ChromeOptions options = new ChromeOptions();
			//options.addArguments("user-data-dir=C:/Users/FB/AppData/Local/Google/Chrome/User Data");
			//options.addArguments("--profile-directory=Profile 6");
			//options.addArguments("start-maximized");
			//driver = new ChromeDriver(options);
			driver =new ChromeDriver();
			driver.get("https://www.google.com/search?tbs=lf:1,lf_ui:2&tbm=lcl&sxsrf=ALiCzsaF5iuoAtWU09kAy1DuqmkhbE4pqA:1669726929868&q=software+company+in+madurai&rflfq=1&num=10&sa=X&ved=2ahUKEwi1_prEudP7AhWrR2wGHRg2BUEQjGp6BAgUEAE&biw=1091&bih=937&dpr=1#rlfi=hd:;si:;mv:[[10.06597871221181,78.27187456423341],[9.801525629671396,78.00785936647951],null,[9.933778892307512,78.13986696535646],12];start:40");
			Thread.sleep(5000);
			JavascriptExecutor js = (JavascriptExecutor) driver;

			for (;;) {
				driver.findElement(By.xpath("//a[@aria-label=  'Page " + i + "']")).click();
				js.executeScript("window.scrollBy(0,document.body.scrollHeight)");
				Thread.sleep(5000);
				List<WebElement> list1 = driver.findElements(By.xpath("//a[@class='yYlJEf Q7PwXb L48Cpd']"));
				XSSFRow row = sheet.createRow(rows++);
				XSSFCell cell = row.createCell(0);
				

				
				Thread.sleep(5000);
				for (WebElement company1 : list1) {
					String name = company1.getAttribute("href");
					System.out.println(name);
					cell.setCellValue(name);
				}
				i++;
				
			}
			
			
		} catch (org.openqa.selenium.StaleElementReferenceException ex) {

		}
		
	}

}

