package com.pom.automation.ExcelHandle;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.Iterator;

public class KeywordSearchScript {

    public static void main(String[] args) {
    
        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Ayon\\Downloads\\crm\\chromedriver.exe");
        
      
        
        
        DayOfWeek dayOfWeek = LocalDate.now().getDayOfWeek();

        
        String excelFilePath = "C:\\Users\\Ayon\\Downloads\\Excel.xlsx";
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(excelFilePath)) {

            
            Sheet worksheet = workbook.getSheet(dayOfWeek.toString());

            // Initialize the Chrome driver
            //WebDriver driver = new ChromeDriver();
            
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--remote-allow-origins=*");
            WebDriver driver = new ChromeDriver(options);

            // Iterate over the rows in the worksheet starting from the second row (assuming headers are in the first row)
            Iterator<Row> rowIterator = worksheet.iterator();
            rowIterator.next(); // Skip the header row
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                String keyword = row.getCell(0).getStringCellValue();

                // Search the keyword in Google
                driver.get("https://www.google.com");
                WebElement searchBox = driver.findElement(By.name("q"));
                searchBox.sendKeys(keyword);
                searchBox.sendKeys(Keys.RETURN);

                // Find the longest and shortest options
                WebElement firstResult = driver.findElement(By.cssSelector("div.g:first-child h3"));
                WebElement lastResult = driver.findElement(By.cssSelector("div.g:last-child h3"));

                String longestOption = firstResult.getText();
                String shortestOption = lastResult.getText();

                // Write the longest and shortest options to the respective cells in the Excel file
                Cell longestCell = row.createCell(1);
                longestCell.setCellValue(longestOption);
                Cell shortestCell = row.createCell(2);
                shortestCell.setCellValue(shortestOption);
            }

            // Save the changes to the Excel file
            workbook.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

