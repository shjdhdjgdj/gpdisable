package com.mainMethod.automation;

import java.io.File;
import java.io.FileNotFoundException;
import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import config.VARIABLES;

public class PageBean {

    private WebDriver driver;
    private WebDriverWait wait;

    public PageBean(WebDriver driver) {
        this.driver = driver;
        this.wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        PageFactory.initElements(driver, this);
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  UTILITY
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Clears Bengali digits / stale values and types the correct Excel value.
     * Handles StaleElementReferenceException with one retry.
     */
    public void clearAndType(WebElement element, String value) {
        try {
            wait.until(ExpectedConditions.elementToBeClickable(element));
            element.click();
            element.clear();
            element.sendKeys(Keys.CONTROL + "a");
            element.sendKeys(Keys.DELETE);
            element.sendKeys(value);
        } catch (StaleElementReferenceException e) {
            try {
                Thread.sleep(300);
                element.clear();
                element.sendKeys(Keys.CONTROL + "a");
                element.sendKeys(Keys.DELETE);
                element.sendKeys(value);
            } catch (Exception ignored) {}
        }
    }

    /**
     * Helper: restore implicit wait to default after any short-wait block.
     * Always call this after temporarily lowering the implicit wait.
     */
    private void restoreImplicitWait() {
        // FIX #1: logicToSkip() was lowering implicit wait to 1s and never restoring it.
        // Every element lookup for the rest of the test was racing against 1s instead of 10s.
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  LOGIN PAGE ELEMENTS
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "inputUserName")
    private WebElement email;

    @FindBy(id = "inputPassword")
    private WebElement passWord;

    @FindBy(id = "insurance_user_season")
    private WebElement seasonDropdown;

    @FindBy(id = "user_session")
    private WebElement sessionDropdown;

    @FindBy(xpath = "//input[@type='checkbox']")
    private WebElement checkbox;

    @FindBy(id = "generate_otp")
    private WebElement generateOtp;

    @FindBy(xpath = "//button[@class=\"btn btn-group btn-default btn-animated btn_login\"]")
    private WebElement loginButton;

    @FindBy(id = "otp")
    private WebElement otp;

    // ─────────────────────────────────────────────────────────────────────────
    //  NAVIGATION
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(xpath = "//*[@id='navbar-collapse-1']/div[2]/ul/li[3]/a")
    private WebElement listToGo;

    @FindBy(xpath = "//*[@id=\"navbar-collapse-1\"]/div[2]/ul/li[3]/ul/li[1]/a")
    private WebElement list;

    // ─────────────────────────────────────────────────────────────────────────
    //  SEARCH SECTION
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "insure_voter_id")
    private WebElement voterCardNumber;

    @FindBy(id = "insur_search")
    private WebElement searchButton;

    @FindBy(xpath = "//tbody[@id='tbodycrop']/tr/td[2]")
    private WebElement already_existing_crop;

    @FindBy(xpath = "//tbody[@id='tbodycrop']/tr/td[5]")
    private WebElement already_existing_gram_panchayat;

    @FindBy(id = "insure_aadhar_no")
    private WebElement aadharCardNumber;

    @FindBy(id = "insure_app_type")
    private WebElement applicationSource;

    // ─────────────────────────────────────────────────────────────────────────
    //  FARMER DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "insure_name")
    private WebElement nameAsPerEpic;

    @FindBy(id = "insure_f_name")
    private WebElement fatherOrHusbandName;

    @FindBy(id = "insure_f_relation")
    private WebElement relationWithFarmerDropDown;

    @FindBy(id = "insure_age")
    private WebElement ageDropDown;

    @FindBy(id = "insure_gender")
    private WebElement genderDropDown;

    @FindBy(id = "insure_caste")
    private WebElement casteDropDown;

    @FindBy(id = "insure_mobile_no")
    private WebElement mobileNumber;

    @FindBy(id = "insure_f_category")
    private WebElement farmerCategoryDropDown;

    @FindBy(id = "insure_nominee_name")
    private WebElement nomineeName;

    @FindBy(id = "insure_id_proof")
    private WebElement voterIDUpload;

    @FindBy(id = "insure_aadhar_doc")
    private WebElement aadharIDUpload;

    // ─────────────────────────────────────────────────────────────────────────
    //  RESIDENTIAL ADDRESS
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "f_district")
    private WebElement farmersResidentialAddressDistrictDropDown;

    @FindBy(id = "block_id")
    private WebElement farmersResidentialAddressblockDropDown;

    @FindBy(id = "gp_id")
    private WebElement farmersResidentialAddressgramPanchayatDropDown;

    @FindBy(id = "vill_id")
    private WebElement farmersResidentialAddressvillageDropDown;

    @FindBy(id = "pin_code")
    private WebElement pinCode;

    // ─────────────────────────────────────────────────────────────────────────
    //  CROP DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_district_id")
    private WebElement cropDetailsDistrictDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_block_id")
    private WebElement cropDetailsBlockDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_crop_id")
    private WebElement cropDetailsCropDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_gram_panchayat_id")
    private WebElement cropDetailsGramPanchayatInitial;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_gram_panchayat_id")
    private WebElement cropDetailsGramPanchayatFinal;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_mouza_id")
    private WebElement cropDetailsMouzaDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_khatian_no")
    private WebElement cropDetailskhaitanNumber;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_plot_no")
    private WebElement cropDetailsPlotNumber;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_inc_land_in_acer")
    private WebElement cropDetailsAreaInAcre;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_area_insured")
    private WebElement cropDetailsAreaInDecimal;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_nature_of_farmer")
    private WebElement cropDetailsNatureOfFarmerDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_form18_document")
    private WebElement cropDetailsNonOwnerCultivatorCertificateUpload;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_insurance_lands_attributes_0_parcha_document")
    private WebElement cropDetailsParchaUpload;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_land_document")
    private WebElement landDocumentProofUpload;

    // ─────────────────────────────────────────────────────────────────────────
    //  BANK DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_account_holder_name")
    private WebElement bankDetailsAccountHolderName;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_account_number")
    private WebElement bankDetailsAccountNumber;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_account_type")
    private WebElement accountTypeDropDown;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_account_ifsc")
    private WebElement ifsCode;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_bank_name")
    private WebElement bankName;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_branch_name")
    private WebElement branchName;

    @FindBy(id = "insurance_farmer_insurance_applications_attributes_0_bank_document")
    private WebElement bankDocumentProofUpload;

    @FindBy(id = "before_insure_submit")
    private WebElement submitButton;

    // ─────────────────────────────────────────────────────────────────────────
    //  NAVIGATION METHOD
    // ─────────────────────────────────────────────────────────────────────────

    public void gotoPage() {
        Actions action = new Actions(driver);
        action.moveToElement(listToGo).perform();
        list.click();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  LOGIN
    // ─────────────────────────────────────────────────────────────────────────

    public void login(String s1, String s2, int index1, int index2) throws InterruptedException {
        email.sendKeys(s1);
        passWord.sendKeys(s2);

        Select dropdown1 = new Select(seasonDropdown);
        dropdown1.selectByIndex(index2);

        wait.until(driver1 -> {
            Select dropDown2 = new Select(sessionDropdown);
            return dropDown2.getOptions().size() > 1;
        });
        new Select(sessionDropdown).selectByIndex(index1);

        Thread.sleep(2000);
        generateOtp.click();
        Thread.sleep(30000);

        if (!checkbox.isSelected()) {
            checkbox.click();
        }
        loginButton.click();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  SEARCH
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * FIX #2: Old code clicked search TWICE in a loop.
     * The second click fired before the first result loaded, leaving the
     * previous row's result in the DOM → logicToSkip() matched stale data → SKIP.
     *
     * Now: search once, then wait 1.5 s for the DOM to settle before
     * logicToSkip() is called by the caller.
     */
    public void searchPerson(String voterCard) {
        wait.until(ExpectedConditions.presenceOfElementLocated(By.id("insure_voter_id")));

        // Uncheck if checked — only once, not in a loop
        if (checkbox.isSelected()) {
            checkbox.click();
        }

        voterCardNumber.clear();
        voterCardNumber.sendKeys(voterCard);
        searchButton.click();

        // Allow the server response / DOM update to settle
        try { Thread.sleep(1500); } catch (InterruptedException ignored) {}
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  SKIP LOGIC
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * FIX #1 (CRITICAL): The implicit wait was lowered to 1 s inside this method
     * but NEVER restored. Every element lookup for the rest of the test was
     * racing against 1 s instead of 10 s, causing failures well before bank details.
     *
     * Fix: restore implicit wait in a finally block.
     */
    public boolean logicToSkip(String crop, String gramPanchayat) {
        try {
            driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(1));
            if (already_existing_crop.isDisplayed() && already_existing_gram_panchayat.isDisplayed()) {
                String existingCrop = already_existing_crop.getText().trim();
                String existingGP   = already_existing_gram_panchayat.getText().trim();
                if (existingCrop.equals(crop.trim()) && existingGP.equals(gramPanchayat.trim())) {
                    System.out.println("⚠️  Skip match — existing crop: [" + existingCrop
                            + "] GP: [" + existingGP + "]");
                    return true;
                }
            }
        } catch (Exception ignored) {
            // Element not present → no existing record → do NOT skip
        } finally {
            restoreImplicitWait(); // ✅ Always restore — was missing before
        }
        return false;
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  AADHAR / APPLICATION TYPE
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * FIX #3: Old code used getText() on an <input> element, which always
     * returns "". Must use getAttribute("value") to read an input's current value.
     */
    public void dataEntry(String aadharCard) {
        String currentValue = aadharCardNumber.getAttribute("value"); // ✅ fixed
        if (currentValue == null || currentValue.trim().isEmpty()) {
            aadharCardNumber.sendKeys(aadharCard);
        }
        wait.until(ExpectedConditions.elementToBeClickable(applicationSource));
        new Select(applicationSource).selectByIndex(1);
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  FARMER DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    public void farmerDetails(String name, String fatherHusbandName, String relationWithFarmer,
            String age, String gender, String caste, String mobileNum,
            String farmerCategory, String epicIDImage, String aadharImg) {

        wait.until(ExpectedConditions.elementToBeClickable(nameAsPerEpic));
        nameAsPerEpic.sendKeys(name);
        fatherOrHusbandName.sendKeys(fatherHusbandName);

        new Select(relationWithFarmerDropDown).selectByValue(relationWithFarmer);
        new Select(ageDropDown).selectByValue(age);
        new Select(genderDropDown).selectByValue(gender);
        new Select(casteDropDown).selectByValue(caste);

        mobileNumber.sendKeys(mobileNum);
        new Select(farmerCategoryDropDown).selectByValue(farmerCategory);

        // Clear auto-filled nominee name
        try {
            Thread.sleep(1000);
            nomineeName.clear();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        // Upload Voter / Epic ID document
        uploadFile(voterIDUpload, VARIABLES.VOTER_FILE_PATH, epicIDImage, "Voter ID");

        // Upload Aadhar document (filename = aadhar number)
        uploadFile(aadharIDUpload, VARIABLES.AADHAR_FILE_PATH, aadharImg, "Aadhar");
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  RESIDENTIAL ADDRESS
    // ─────────────────────────────────────────────────────────────────────────

    public void farmerResidentialAddress(String district, String block, String gramPanchayat,
            String village, String pin) throws InterruptedException {

        new Select(farmersResidentialAddressDistrictDropDown).selectByVisibleText(district);

        wait.until(d -> new Select(farmersResidentialAddressblockDropDown).getOptions().size() > 1);
        new Select(farmersResidentialAddressblockDropDown).selectByVisibleText(block);

        wait.until(d -> new Select(farmersResidentialAddressgramPanchayatDropDown).getOptions().size() > 1);
        new Select(farmersResidentialAddressgramPanchayatDropDown).selectByVisibleText(gramPanchayat);

        wait.until(d -> new Select(farmersResidentialAddressvillageDropDown).getOptions().size() > 1);
        Thread.sleep(2000);
        new Select(farmersResidentialAddressvillageDropDown).selectByIndex(1);

        pinCode.sendKeys(pin);
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  CROP DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    public void cropDetailsEntry(String district, String block, String crop, String gpInitial,
            String mouza, String khatianNumber, String plotNumber,
            String areaInAcre1, String natureOfFarmer, String parchaImg)
            throws InterruptedException {

        new Select(cropDetailsDistrictDropDown).selectByVisibleText(district);

        wait.until(d -> new Select(cropDetailsBlockDropDown).getOptions().size() > 1);
        new Select(cropDetailsBlockDropDown).selectByVisibleText(block);

        wait.until(d -> new Select(cropDetailsCropDropDown).getOptions().size() > 1);
        new Select(cropDetailsCropDropDown).selectByVisibleText(crop);

        // GP Initial dropdown (shown for some crops)
        if (cropDetailsGramPanchayatInitial.isEnabled()) {
            wait.until(d -> new Select(cropDetailsGramPanchayatInitial).getOptions().size() > 1);
            new Select(cropDetailsGramPanchayatInitial).selectByVisibleText(gpInitial);
        }

        // GP Final dropdown (shown for other crops)
        if (cropDetailsGramPanchayatFinal.isEnabled()) {
            wait.until(d -> new Select(cropDetailsGramPanchayatFinal).getOptions().size() > 1);
            new Select(cropDetailsGramPanchayatFinal).selectByVisibleText(gpInitial);
        }

        wait.until(d -> new Select(cropDetailsMouzaDropDown).getOptions().size() > 1);
        new Select(cropDetailsMouzaDropDown).selectByVisibleText(mouza);

        // Double-clear khatian & plot to remove Bengali / stale content
        clearAndType(cropDetailskhaitanNumber, khatianNumber);
        clearAndType(cropDetailsPlotNumber, plotNumber);
        Thread.sleep(500);
        cropDetailskhaitanNumber.clear();
        cropDetailsPlotNumber.clear();
        clearAndType(cropDetailskhaitanNumber, khatianNumber);
        clearAndType(cropDetailsPlotNumber, plotNumber);

        cropDetailsAreaInAcre.sendKeys(areaInAcre1);

        // Click nature-of-farmer dropdown to trigger any area-validation alert
        cropDetailsNatureOfFarmerDropDown.click();

        double area = Double.parseDouble(areaInAcre1);
        if (area >= 1) {
            wait.until(ExpectedConditions.alertIsPresent());
            driver.switchTo().alert().accept();
        }

        wait.until(d -> new Select(cropDetailsNatureOfFarmerDropDown).getOptions().size() > 1);
        new Select(cropDetailsNatureOfFarmerDropDown).selectByVisibleText(natureOfFarmer);

        // Upload parcha document
        uploadFile(cropDetailsParchaUpload, VARIABLES.PARCHA_FILE_PATH, parchaImg, "Parcha");

        // Upload land document (same file as parcha)
        uploadFile(landDocumentProofUpload, VARIABLES.PARCHA_FILE_PATH, parchaImg, "Land document");
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  BANK DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * FIX #4 (CRITICAL): Old code used selectByValue(accountType) where accountType
     * was the Excel visible text e.g. "Savings" / "Current".
     * selectByValue() matches the HTML <option value="..."> attribute, NOT the
     * visible label → throws "Cannot locate option with value: Savings".
     *
     * Fix: use selectByVisibleText() so it matches the label the user sees.
     */
    public void bankDetailsEntry(String name, String accountNumber, String accountType,
            String ifscCode) throws InterruptedException {

        // 1. Account Holder Name
        if (bankDetailsAccountHolderName.isEnabled()
                && bankDetailsAccountHolderName.getAttribute("readonly") == null) {
            bankDetailsAccountHolderName.sendKeys(name);
        }

        // 2. Account Number
        if (bankDetailsAccountNumber.isEnabled()
                && bankDetailsAccountNumber.getAttribute("readonly") == null) {
            bankDetailsAccountNumber.clear();
            bankDetailsAccountNumber.sendKeys(accountNumber);
        }

        // 3. Account Type — FIX: selectByVisibleText, NOT selectByValue
        if (accountTypeDropDown.isEnabled()) {
            try {
                new Select(accountTypeDropDown).selectByVisibleText(accountType); // ✅ fixed
            } catch (NoSuchElementException e) {
                // Fallback: try case-insensitive match by iterating options
                Select sel = new Select(accountTypeDropDown);
                boolean matched = false;
                for (WebElement opt : sel.getOptions()) {
                    if (opt.getText().trim().equalsIgnoreCase(accountType.trim())) {
                        opt.click();
                        matched = true;
                        break;
                    }
                }
                if (!matched) {
                    System.err.println("⚠️  Account type option not found: [" + accountType
                            + "] — selecting index 1 as fallback");
                    sel.selectByIndex(1);
                }
            }
        }

        // 4. IFSC Code → triggers bank name auto-fill
        if (ifsCode.isEnabled() && ifsCode.getAttribute("readonly") == null) {
            ifsCode.clear();
            ifsCode.sendKeys(ifscCode);
            if (bankName.isEnabled()) {
                bankName.click();
                Thread.sleep(500);
            }
        }

        // 5. Bank document upload (filename = account number)
        if (bankDocumentProofUpload.isEnabled()
                && bankDocumentProofUpload.getAttribute("readonly") == null) {
            uploadFile(bankDocumentProofUpload, VARIABLES.BANK_FILE_PATH, accountNumber, "Bank document");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  SUBMIT
    // ─────────────────────────────────────────────────────────────────────────

    public void submitForm() {
        submitButton.click();
    }

    // ─────────────────────────────────────────────────────────────────────────
    //  PRIVATE HELPERS
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Tries jpg → jpeg → png → pdf for the given filename under basePath.
     * Sends the absolute path to the upload input element.
     * Logs a clear error if no file is found (does NOT throw, so the test
     * continues and can be marked FAIL by the caller if needed).
     */
    private void uploadFile(WebElement uploadElement, String basePath, String fileName, String label) {
        String[] extensions = { ".jpg", ".jpeg", ".png", ".pdf" };
        for (String ext : extensions) {
            File f = new File(basePath + "\\" + fileName + ext);
            if (f.exists()) {
                uploadElement.sendKeys(f.getAbsolutePath());
                return;
            }
        }
        // File not found — log but don't crash; caller decides how to handle
        System.err.println("⚠️  " + label + " file not found for: " + fileName
                + "  (looked in: " + basePath + ")");
    }
}