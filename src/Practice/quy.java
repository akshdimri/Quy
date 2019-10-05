package Practice;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.sikuli.script.FindFailed;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class quy {
	public static int i;

	public static void main(String[] args)throws InterruptedException, FindFailed, BiffException, IOException  {
		
			
			System.setProperty("webdriver.gecko.driver", "C:\\Users\\admin\\Downloads\\ECLIPSE\\geckodriver-v0.24.0-win64\\geckodriver.exe");
				WebDriver driver= new FirefoxDriver();

				driver.navigate().to("http://www.quytech.com/");
				//Thread.sleep(3000);
		
				//Thread.sleep(1000);
		//driver.findElement(By.xpath("(//*[@class='btn btn-default btn-lg orange_butn mrgn-rgt-lg header_btn'])[1] ")).click();
		
		
		
	
				
				File f1 = new File("C:\\Users\\admin\\Desktop\\test - Copy (3).xls");
				Workbook rwb = Workbook.getWorkbook(f1);
				Sheet rsh = rwb.getSheet("Sheet1");

				WritableWorkbook wwb = Workbook.createWorkbook(f1, rwb);
				WritableSheet wsh = wwb.getSheet(1);

				int nur = rsh.getRows();
				int nuc = rsh.getColumns();
				System.out.println("No. of used Rows  are :" + nur);
				System.out.println("No. of used Columns  are :" + nuc);
				
				
				for (int i=1;i<nur;i++) {
					
					
					JavascriptExecutor js=(JavascriptExecutor) driver;
					js.executeScript("window.scrollBy(0,5500)");
			
					
					String Name = rsh.getCell(1, i).getContents();
					String email = rsh.getCell(2, i).getContents();
					String num = rsh.getCell(3, i).getContents();
					String sub = rsh.getCell(4, i).getContents();
					String msg = rsh.getCell(5, i).getContents();		
					String crt = rsh.getCell(6, i).getContents();
					
					
					System.out.println(Name);
					System.out.println(email);
					System.out.println(num);
					System.out.println(sub);
					System.out.println(msg);
					System.out.println(crt);
					
					
					
					driver.findElement(By.xpath("(//*[@class='contact_form'])[1]")).sendKeys(Name);
					driver.findElement(By.xpath("(//*[@class='contact_form'])[2]")).sendKeys(email);
					driver.findElement(By.xpath("(//*[@class='contact_form'])[3]")).sendKeys(num);
					driver.findElement(By.xpath("(//*[@class='contact_form'])[4]")).sendKeys(sub);
					driver.findElement(By.xpath("(//*[@class='contact_form'])[5]")).sendKeys(msg);
				
					
					Thread.sleep(7000);
					
					
          driver.findElement(By.xpath("//*[@id='submit_form']")).click();
				
			if(crt.equalsIgnoreCase("invalid"))
{
				driver.navigate().refresh();
				
				}
				
				
				}
				
				
				
					
					
				//driver.close();
					
					
					
					}
				
				
	
		
		
		
		
			
		
		

		
	}


