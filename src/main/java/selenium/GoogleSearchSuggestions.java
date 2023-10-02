package selenium;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;

public class GoogleSearchSuggestions {
    public static void main(String[] args) {
        // Get the current day of the week (e.g., "Thursday")
        SimpleDateFormat dateFormat = new SimpleDateFormat("EEEE");
        String currentDay = dateFormat.format(new Date());

        // Initialize WebDriver (Firefox)
        WebDriver driver = new FirefoxDriver();

        // Load the existing Excel file
        String excelFilePath = "/Users/badrulalam/eclipse-workspace/seleniumproject/Excel.xlsx";
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Check if the current day's sheet exists in the workbook
            if (workbook.getSheetIndex(currentDay) != -1) {
                Sheet worksheet = workbook.getSheet(currentDay);

                // Start from the 3rd row (index 2) and 3rd column (index 2)
                for (int rowIndex = 1; rowIndex < worksheet.getPhysicalNumberOfRows(); rowIndex++) {
                    Row row = worksheet.getRow(rowIndex);

                    // Get the keyword from the 3rd column (index 2)
                    Cell keywordCell = row.getCell(1);
                    if (keywordCell != null) {
                        String keyword = keywordCell.getStringCellValue();
                        if (keyword != null && !keyword.isEmpty()) {
                            // Go to Google
                            driver.get("https://www.google.com");

                            // Find the search input field by name
                            WebElement searchBox = driver.findElement(By.name("q"));

                            // Enter the keyword
                            searchBox.clear();
                            searchBox.sendKeys(keyword);
//                            searchBox.submit();

                            // Wait for the suggestions to appear (maximum 10 seconds)
                            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//ul[@role='listbox']/li[@role='presentation']")));

                            // Get the suggestions from the search bar
                            List<WebElement> suggestions = driver.findElements(By.xpath("//ul[@role='listbox']/li[@role='presentation']"));
                            
                            String maxSuggestion = "";
                            String minSuggestion = suggestions.get(0).getText();

                            // Find the suggestion with the maximum and minimum text length
                            for (WebElement suggestion : suggestions) {
                                String text = suggestion.getText();
                                if (text.length() > maxSuggestion.length()) {
                                    maxSuggestion = text;
                                }
                                if (text.length() < minSuggestion.length()) {
                                    minSuggestion = text;
                                }
                            }


                            // Update the corresponding columns in the Excel file
                            row.createCell(2).setCellValue(minSuggestion);
                            row.createCell(3).setCellValue(maxSuggestion);
                        }
                    }
                }

                // Save the updated Excel file
                try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                    workbook.write(fos);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Close the browser
            driver.quit();
        }
    }
}
