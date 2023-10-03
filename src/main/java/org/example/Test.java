package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.*;
import java.util.ArrayList;
import java.util.List;


public class Test {

    public static void main(String[] args) throws IOException, InvalidFormatException, InterruptedException {
        String filepath = "D:\\SQA\\Excel.xlsx";
        File myfile = new File(filepath);


        FileInputStream fis = new FileInputStream(myfile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Saturday");
        int rowcount = sheet.getLastRowNum();
        int colcount = sheet.getRow(3).getLastCellNum();

//        System.out.println("rowcount :" + rowcount + " " + "colcount :" + colcount);

        ChromeDriver driver = new ChromeDriver();
        driver.get("https://www.google.com");
        Thread.sleep(2000);

        for (int i = 2; i <= rowcount; i++) {
            XSSFRow celldata = sheet.getRow(i);
            String keyword = celldata.getCell(2).getStringCellValue();


            driver.findElement(By.name("q")).sendKeys(keyword);
            Thread.sleep(10000);

//            Dynamic keyword search
//            List<WebElement> list = driver.findElements(By.cssSelector("div[aria-label*='"+keyword.toLowerCase()+"'] div:nth-child(1) span"));
            List<WebElement> list = driver.findElements(By.xpath("//ul[@role='listbox']/li/descendant::div[@class='wM6W7d']"));

//            System.out.println(list.size());

            ArrayList<String> searchData = new ArrayList<String>();

            for (int n = 0; n < list.size(); n++) {
                searchData.add(list.get(n).getText());
                System.out.println(list.get(n).getText());
            }
            String minString = searchData.get(0);
            String maxString = searchData.get(0);

            for (String str : searchData) {
                if (str.length() < minString.length()) {
                    minString = str;
                }
                if (str.length() > maxString.length()) {
                    maxString = str;
                }
            }

//            System.out.println("Largest: "+maxString);
//            System.out.println("Smallest: " +minString);


            Cell cell = celldata.createCell(3);
            cell.setCellValue(maxString);
            Cell cell1 = celldata.createCell(4);
            cell1.setCellValue(minString);

            FileOutputStream fos = new FileOutputStream(filepath);
            workbook.write(fos);
            fos.close();


            driver.findElement(By.name("q")).clear();
            searchData.clear();



        }


    }


}
