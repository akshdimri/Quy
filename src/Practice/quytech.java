package Practice;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Screen;
import org.sikuli.script.ScreenImage;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class quytech {
	
	public static int i;
	
	public static void main(String args[]) throws InterruptedException, FindFailed, BiffException, IOException {
		System.setProperty("webdriver.gecko.driver", "C:\\Users\\admin\\Downloads\\ECLIPSE\\geckodriver-v0.24.0-win64\\geckodriver.exe");
		WebDriver driver= new FirefoxDriver();

		driver.navigate().to("http://www.quytech.com/");
		Thread.sleep(3000);
		
		driver.findElement(By.xpath("(//*[@class='fa fa-angle-right'])[1]")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("(//*[@class='fa fa-angle-right'])[1]")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("(//*[@class='fa fa-angle-right'])[1]")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("(//*[@class='fa fa-angle-right'])[1]")).click();
		
		Thread.sleep(1000);
		
		
		driver.findElement(By.xpath("//*[@id=\"js-navbar-collapse\"]/ul/li[1]/a")).click();
		driver.findElement(By.xpath("(//*[@class='dropdown-toggle'])[1]")).click();
		driver.findElement(By.xpath("(//*[@href='http://www.quytech.com/company-overview.php'])[1]")).click();
		
		
		JavascriptExecutor js=(JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,1200)");
		

		for(i=1;i<=9;i++) {
			
			
			driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])"+"["+i+"]")).click();
			driver.navigate().back();
			
			
			
			Thread.sleep(1000);
			
			
		}
		
		
		
	/*driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[1]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[2]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[3]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[4]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[5]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[6]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[7]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[8]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
		driver.findElement(By.xpath("(//*[@class='pull-left services_icon'])[9]")).click();
		driver.navigate().back();
		js.executeScript("window.scrollBy(0,500)");
	*/	
	
		driver.findElement(By.xpath("//*[@class='zsiq_user siqico-chat']")).click();
		Thread.sleep(1000);
		
		
		//driver.findElement(By.xpath("//*[@id='visname']")).sendKeys("");
		
		Screen s=new Screen();
		
		//s.click("Capture.png");
		
		s.type("Capture.png","Manoj");
		
		s.type("email.png","manoj.yadav@quytech.com");
		
		s.type("msg.png","This is a test");
		
		
		
		s.click("enter.png");
		
		s.click("arrow.png");
		
		s.click("Minimize.png");

		
		driver.close();
		
		
	
		File f1 = new File("C:\\Users\\admin\\Desktop\\test.xls\\test.xls");
		Workbook rwb = Workbook.getWorkbook(f1);
		Sheet rsh = rwb.getSheet("Sheet1");

		WritableWorkbook wwb = Workbook.createWorkbook(f1, rwb);
		WritableSheet wsh = wwb.getSheet(1);

		int nur = rsh.getRows();
		int nuc = rsh.getColumns();
		System.out.println("No. of used Rows  are :" + nur);
		System.out.println("No. of used Columns  are :" + nuc);
		
		
		
		
		
		
		//error xpath://*[@id='email-error']
		
		
		
		
		
	}
}
