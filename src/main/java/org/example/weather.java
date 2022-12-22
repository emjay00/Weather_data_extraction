//weather data extraction

package org.example;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class weather {
    public static void main(String[] args) throws InterruptedException, IOException {
        System.setProperty("webdriver.chrome.driver", "C:\\Selenium\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://weather.com/");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        WebElement popupExitBtn = driver.findElement(By.cssSelector("path[d='M19.707 19.293l-1.414 1.414L12 14.414l-6.293 6.293-1.414-1.414L10.586 13 4.293 6.707l1.414-1.414L12 11.586l6.293-6.293 1.414 1.414L13.414 13z']"));
        popupExitBtn.click();
        WebElement searchBox = driver.findElement(By.cssSelector("#LocationSearch_input"));
        searchBox.click();
        //Enter location to gather results from
        searchBox.sendKeys("waterloo,on");

        driver.findElement(By.xpath("//*[@id=\"LocationSearch_listbox-0\"]")).click();
        driver.findElement(By.xpath("//span[normalize-space()='Monthly']")).click();
        XSSFWorkbook wf = new XSSFWorkbook();
        XSSFSheet sheet = wf.createSheet();
        sheet.createRow(0);
        String h=driver.findElement(By.xpath("//strong")).getText();
        String place=driver.findElement(By.xpath("//span[@data-testid=\"PresentationName\"]")).getText();
        sheet.getRow(0).createCell(0).setCellValue(h);
        sheet.getRow(0).createCell(1).setCellValue("Location: "+place);
        sheet.createRow(1);
        sheet.getRow(1).createCell(0).setCellValue("Date");
        sheet.getRow(1).createCell(1).setCellValue("Maximum");
        sheet.getRow(1).createCell(2).setCellValue("Minimum");
        for (int i = 2; i <= 36; i++) {
            WebElement data = driver.findElement(By.xpath("//button[contains(@data-id,\"calendar\")][" + (i-1) + "]"));
            WebElement max = driver.findElement(By.xpath("(//div[contains(@class,\"CalendarDateCell--tempHigh\")]/span)[" + (i-1) + "]"));
            WebElement min = driver.findElement(By.xpath("(//div[contains(@class,\"CalendarDateCell--tempLow\")]/span)[" + (i-1) + "]"));
            String date = data.getAttribute("data-id");
            sheet.createRow(i);
            sheet.getRow(i).createCell(0).setCellValue(date);
            String maximum = max.getText();
            sheet.getRow(i).createCell(1).setCellValue(maximum);
            String minimum = min.getText();
            sheet.getRow(i).createCell(2).setCellValue(minimum);
        }
        File file = new File((System.getProperty("user.dir") + "/src/main/resources/WeatherData.xlsx"));
        FileOutputStream fs = new FileOutputStream(file);
        wf.write(fs);
        wf.close();
        
    }
}
