package selenium;

import io.github.bonigarcia.wdm.WebDriverManager;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import.XSSFWorkbook;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

public class Outdoor_toys {
    private WebDriverWait wait;
    private Workbook workbook;
    private Sheet sheet;
    private WebDriver driver;

    @BeforeClass
    @Parameters("browser")
    public void setUp(String browser) {
        if (browser.equalsIgnoreCase("chrome")) {
           // WebDriverManager.chromedriver().setup();
            driver = new ChromeDriver();
        } else if (browser.equalsIgnoreCase("edge")) {
           // WebDriverManager.edgedriver().setup();
            driver = new EdgeDriver();
        }
        wait = new WebDriverWait(driver, Duration.ofSeconds(180));
        workbook = new XSSFWorkbook(); // Ensure this line is present
        sheet = workbook.createSheet("Outdoor Toys");

        // Create header row
        Row headerRow = sheet.createRow(0);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("Toy Name");
        Cell headerCell2 = headerRow.createCell(1);
        headerCell2.setCellValue("Toy URL");
    
    }

    @Test
    public void testOutdoorToysSearch() {
        // Open eBay and maximize the window
        driver.get("https://www.ebay.com/");
        driver.manage().window().maximize();

        // Click on 'Advanced' search
       WebElement advancedSearch = wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Advanced")));
        advancedSearch.click();

        // Enter "Outdoor toys" in the search box
        WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("_nkw")));
        searchBox.sendKeys("Outdoor toys");

        // Select "Any words, any order" in the search type dropdown
        Select searchType = new Select(driver.findElement(By.id("s0-1-17-4[0]-7[1]-_in_kw")));
        searchType.selectByVisibleText("Any words, any order");

        // Select "Toys & Hobbies" category
        Select category = new Select(driver.findElement(By.id("s0-1-17-4[0]-7[3]-_sacat")));
        category.selectByVisibleText("Toys & Hobbies");

        // Check "Title and description"
        WebElement titleDescription = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label[for='s0-1-17-5[1]-[0]-LH_TitleDesc']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", titleDescription);

        // Check "New" condition
        WebElement newCondition = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label[for='s0-1-17-6[4]-[0]-LH_ItemCondition']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", newCondition);

        // Select "Free returns" and "Returns accepted"
        WebElement freeReturns = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label[for='s0-1-17-5[5]-[0]-LH_FR']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", freeReturns);

        WebElement returnsAccepted = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label[for='s0-1-17-5[5]-[1]-LH_RPA']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", returnsAccepted);

        // Set location to "Worldwide"
        WebElement worldwide = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("label[for='s0-1-17-6[7]-[3]-LH_PrefLoc']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", worldwide);

        // Click the search button
        WebElement searchButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[@class='btn btn--primary' and @type='submit']")));
        ((JavascriptExecutor) driver).executeScript("arguments[0].click();", searchButton);

        // Wait for search results to load
        wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector(".s-item__link")));

        // Extract product links from the search results
        List<WebElement> productLinks = driver.findElements(By.cssSelector(".s-item__link"));
        System.out.println("Outdoor Toy Listings:");

        int rowNum = 1;
        for (WebElement link : productLinks) {
            String toyName = link.getText();
            String toyURL = link.getAttribute("href");

            // Verify if the name contains 'toys'
            if (toyName.toLowerCase().contains("toys")) {
                System.out.println(toyName + " -> " + toyURL);

                // Write to Excel sheet
                Row row = sheet.createRow(rowNum++);
                Cell cell1 = row.createCell(0);
                cell1.setCellValue(toyName);
                Cell cell2 = row.createCell(1);
                cell2.setCellValue(toyURL);
            }
        }
    }

    @AfterClass
    public void tearDown() {
        // Save the Excel file
        try (FileOutputStream fileOut = new FileOutputStream("Outdoor_Toys.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the browser
        //driver.quit();
    }
}