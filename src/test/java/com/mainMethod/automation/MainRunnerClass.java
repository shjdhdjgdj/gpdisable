package com.mainMethod.automation;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.NoSuchElementException;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.SkipException;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import Freelance.com.projectSetup.ExcelUtility;
import config.VARIABLES;
import io.github.bonigarcia.wdm.WebDriverManager;

public class MainRunnerClass {

    private WebDriver driver;
    static WebDriverWait wait;
    private PageBean pom;

    // ─────────────────────────────────────────────────────────────────────────
    //  SUITE SETUP
    // ─────────────────────────────────────────────────────────────────────────

    @BeforeSuite
    public void beforeSuite() throws InterruptedException {

        // Load config.properties from the working directory
        Properties prop = new Properties();
        try (FileInputStream fis = new FileInputStream("config.properties")) {
            prop.load(fis);
        } catch (Exception e) {
            throw new RuntimeException("config.properties not found — place it in the run directory", e);
        }

        String browser = prop.getProperty("browser", "chrome").trim();

        if (browser.equalsIgnoreCase("chrome")) {
            WebDriverManager.chromedriver().setup();
            driver = new ChromeDriver();
        } else if (browser.equalsIgnoreCase("firefox")) {
            WebDriverManager.firefoxdriver().setup();
            driver = new FirefoxDriver();
        } else {
            throw new RuntimeException("Unsupported browser in config.properties: " + browser);
        }

        driver.manage().window().maximize();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        pom = new PageBean(driver);

        // Navigate to sign-in page
        try {
            driver.get(VARIABLES.SIGN_IN_PAGE_URL);
        } catch (NoSuchElementException e) {
            checkElementWithRetries(VARIABLES.SIGN_IN_PAGE_URL,
                    "//*[contains(text(),'Insurance Log In')]", 10, 3);
        }

        // Login and navigate to registration page
        try {
            pom.login(VARIABLES.EMAIL, VARIABLES.PASSWORD, 1, 2);
            openNewTab();
            driver.get(VARIABLES.NEW_REGISTRATION_URL);
        } catch (NoSuchElementException | InterruptedException e) {
            e.printStackTrace();
            checkElementWithRetries(VARIABLES.NEW_REGISTRATION_URL,
                    "//h4[contains(text(),'SBI GENERAL INSURANCE COMPANY LIMITED')]", 5, 5);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  DATA PROVIDER
    // ─────────────────────────────────────────────────────────────────────────

    @DataProvider(name = "excelData")
    public Object[][] testMainMethod() {
        return ExcelUtility.getExcelData();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  TEST METHOD
    // ─────────────────────────────────────────────────────────────────────────

    @Test(dataProvider = "excelData")
    public void runTests(Object[] data) throws InterruptedException {

        // data[0] = 1-based Excel row index (inserted by ExcelUtility.getExcelData)
        int rowIndex = (int) data[0];
        String status = "PASS";

        try {
            // ── Extract fields (indices shifted by +1 because data[0] = row index) ──
            //    Adjust column numbers below if your Excel sheet layout differs.
            String FarmrName    = (String) data[3];   // col C
            String FathrHusName = (String) data[4];   // col D
            String EpicID       = (String) data[5];   // col E
            String AadharNo     = (String) data[6];   // col F
            String Age          = (String) data[8];   // col H
            String Gender       = (String) data[9];   // col I
            String Caste        = (String) data[10];  // col J
            String MobNo        = (String) data[11];  // col K
            String Crop         = (String) data[12];  // col L
            String District     = (String) data[13];  // col M
            String Block        = (String) data[14];  // col N
            String GP           = (String) data[15];  // col O
            String Mouza1       = (String) data[16];  // col P
            String KhatianNo1   = (String) data[18];  // col R
            String PlotNo1      = (String) data[19];  // col S
            String AreaInsur1   = (String) data[20];  // col T
            String FarmrCat     = (String) data[21];  // col U
            String NatureFarmr1 = (String) data[22];  // col V
            String IFSCode      = (String) data[23];  // col W
            String AccNo        = (String) data[24];  // col X
            String Vill         = (String) data[25];  // col Y
            String Pin          = (String) data[26];  // col Z
            String AccType      = (String) data[27];  // col AA  e.g. "Savings" / "Current"
            String Relation     = (String) data[28];  // col AB
            String EpicIDImg    = (String) data[30];  // col AD
            String ParchaImg    = (String) data[31];  // col AE

            System.out.println("▶️  Row " + rowIndex + " | Epic: " + EpicID + " | " + FarmrName);

            // ── Verify the registration page is loaded ────────────────────────
            checkElementWithRetries(VARIABLES.NEW_REGISTRATION_URL,
                    "//h4[contains(text(),'SBI GENERAL INSURANCE COMPANY LIMITED')]", 10, 5);

            // ── Search by Epic / Voter ID ─────────────────────────────────────
            pom.searchPerson(EpicID); // FIX #2: now searches only once + waits

            // ── Skip if this crop+GP combo already exists ─────────────────────
            // FIX #1: logicToSkip() now restores implicit wait in finally block
            if (pom.logicToSkip(Crop, GP)) {
                status = "SKIP";
                System.out.println("⏭️  Row " + rowIndex + " — record already exists, skipping");
                throw new SkipException("Record already exists for Crop=[" + Crop + "] GP=[" + GP + "]");
            }

            // ── Fill form sections ────────────────────────────────────────────
            pom.dataEntry(AadharNo);                                        // FIX #3: getText → getAttribute

            pom.farmerDetails(FarmrName, FathrHusName, Relation,
                    Age, Gender, Caste, MobNo, FarmrCat, EpicIDImg, AadharNo);

            pom.farmerResidentialAddress(District, Block, GP, Vill, Pin);

            pom.cropDetailsEntry(District, Block, Crop, GP,
                    Mouza1, KhatianNo1, PlotNo1, AreaInsur1, NatureFarmr1, ParchaImg);

            pom.bankDetailsEntry(FarmrName, AccNo, AccType, IFSCode); // FIX #4: selectByVisibleText

            pom.submitForm();

            System.out.println("✅ Row " + rowIndex + " — submitted successfully");

        } catch (SkipException e) {
            status = "SKIP";
            throw e; // re-throw so TestNG marks the test as skipped

        } catch (Exception e) {
            status = "FAIL";
            System.err.println("❌ Row " + rowIndex + " — FAILED: " + e.getMessage());
            throw e; // re-throw so TestNG marks the test as failed

        } finally {
            // Always write result back to Excel regardless of outcome
            ExcelUtility.updateTestStatus(rowIndex, status);
            System.out.println("📊 Excel updated — Row " + rowIndex + " = " + status);
            System.out.println("──────────────────────────────────────────────────");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  AFTER METHOD — refresh between rows
    // ─────────────────────────────────────────────────────────────────────────

    @AfterMethod
    public void pageRefresh() {
        try {
            driver.navigate().refresh();
            Thread.sleep(2000);
        } catch (Exception e) {
            System.err.println("⚠️  Page refresh failed: " + e.getMessage());
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  AFTER SUITE — quit browser
    // ─────────────────────────────────────────────────────────────────────────

    @AfterSuite
    public void afterSuite() {
        if (driver != null) {
            driver.quit();
            System.out.println("✅ Browser closed. Suite complete.");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  HELPERS
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Retries finding an element on a given URL, opening new tabs as needed.
     * Gives up gracefully after maxTabSwitches attempts.
     */
    public void checkElementWithRetries(String url, String xpath,
            int maxRetries, int maxTabSwitches) throws InterruptedException {

        boolean found    = false;
        int retryCount   = 0;
        int tabCount     = 0;

        while (!found && tabCount < maxTabSwitches) {
            while (retryCount < maxRetries) {
                try {
                    if (driver.findElement(By.xpath(xpath)).isDisplayed()) {
                        found = true;
                        break;
                    } else {
                        driver.navigate().refresh();
                        Thread.sleep(2000);
                    }
                } catch (NoSuchElementException e) {
                    retryCount++;
                    if (retryCount >= maxRetries) {
                        System.out.println("⚠️  Max retries reached — opening new tab");
                        openNewTab();
                        driver.get(url);
                        retryCount = 0;
                        tabCount++;
                        break;
                    }
                }
            }
        }

        if (!found) {
            System.out.println("⚠️  Element not found after " + maxTabSwitches + " tab attempts: " + xpath);
        }
    }

    private void openNewTab() {
        ((JavascriptExecutor) driver).executeScript("window.open('about:blank', '_blank');");
        String current = driver.getWindowHandle();
        for (String handle : driver.getWindowHandles()) {
            if (!handle.equals(current)) {
                driver.switchTo().window(handle);
                break;
            }
        }
    }
}