package Amazon.in;

import java.io.*;
import java.nio.file.*;
import java.text.*;
import java.time.Duration;
import java.util.*;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.*;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;

import javax.activation.*;
import javax.mail.*;
import javax.mail.internet.*;

public class Cab_Data {

    public static void main(String[] args) throws Exception {
        // ======= HEADLESS MODE CONFIGURATION =======
        boolean HEADLESS_MODE = true;

        // ======= Chrome Options Setup =======
        ChromeOptions options = new ChromeOptions();
        String downloadPath = System.getProperty("user.home") + "\\Downloads";
        Map<String, Object> prefs = new HashMap<>();
        prefs.put("profile.default_content_settings.popups", 0);
        prefs.put("download.default_directory", downloadPath);
        prefs.put("download.prompt_for_download", false);
        prefs.put("download.directory_upgrade", true);
        prefs.put("safebrowsing.enabled", true);
        options.setExperimentalOption("prefs", prefs);

        if (HEADLESS_MODE) {
            System.out.println("🔧 Running in HEADLESS mode...");
            options.addArguments("--headless");
            options.addArguments("--window-size=1920,1080");
            options.addArguments("--remote-allow-origins=*");
            options.addArguments("--disable-web-security");
        } else {
            System.out.println("🖥️ Running in GUI mode...");
            options.addArguments("--remote-allow-origins=*");
        }

        WebDriver driver = new ChromeDriver(options);
        if (!HEADLESS_MODE) {
            driver.manage().window().maximize();
        }
        driver.get("https://amz.moveinsync.com/Hyd/");

        // ===== Login Steps =====
        driver.findElement(By.xpath("//input[@class='loginInput']")).sendKeys("euro-ashraf-ops@dummy.com");
        driver.findElement(By.xpath("//input[@class='loginInput passInput']")).sendKeys("Euro@789456");
        driver.findElement(By.xpath("//input[@class='loginBtn loginEnabled']")).click();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

        // ===== Vehicle Report Actions =====
        By refreshLocator = By.xpath("//img[@title='Click to refresh now']");
        WebElement refreshBtn = wait.until(ExpectedConditions.elementToBeClickable(refreshLocator));
        refreshBtn.click();
        wait.until(ExpectedConditions.elementToBeClickable(refreshLocator));

        WebElement SettingsBtn = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='billing']")));
        Actions action = new Actions(driver);
        action.moveToElement(SettingsBtn).perform();

        By manageVehicleLocator = By.xpath("//li[@class='CabManagement roleBasedAccess']");
        WebElement Manage_Vehicle = wait.until(ExpectedConditions.elementToBeClickable(manageVehicleLocator));
        Manage_Vehicle.click();

        wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.xpath("//iframe[@id='setting-dashboards']")));
        By Manage_ComplianceLocator = By.cssSelector(".manage-compliance-link.ng-tns-c2299869914-0");
        WebElement Manage_Compliance = wait.until(ExpectedConditions.visibilityOfElementLocated(Manage_ComplianceLocator));
        String mainWindow = driver.getWindowHandle();
        Manage_Compliance.click();

        wait.until(ExpectedConditions.numberOfWindowsToBe(2));
        Set<String> allWindows = driver.getWindowHandles();
        for (String window : allWindows) {
            if (!window.equals(mainWindow)) {
                driver.switchTo().window(window);
                break;
            }
        }
        System.out.println("✅ Switched to new page: " + driver.getTitle());

        By New_WindowLocator = By.xpath("//mat-accordion[@class=\"mat-accordion\"]/mat-expansion-panel[2]");
        WebElement Active_Vehicle = wait.until(ExpectedConditions.visibilityOfElementLocated(New_WindowLocator));
        Active_Vehicle.click();
        By excelBtnLocator = By.xpath("//img[@alt='Excel']");
        WebElement Vehicle_ExcelBtn = wait.until(ExpectedConditions.elementToBeClickable(excelBtnLocator));
        action.moveToElement(Vehicle_ExcelBtn).click().build().perform();
        WebElement Vehicle_CustomReport = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(text(),'Download Custom Report')]")));
        Vehicle_CustomReport.click();
        Thread.sleep(3000);
        String vehicleFileName = renameLatestFile("Active Vehicle");
        filterExcel(vehicleFileName, getVehicleColumns());

        // ===== Driver Report Actions =====
        By driverLabelLocator = By.xpath("//label[@for='Driver']");
        WebElement driverLabel = wait.until(ExpectedConditions.elementToBeClickable(driverLabelLocator));
        driverLabel.click();

        By ActiveDriverLocator = By.xpath("//mat-expansion-panel-header[@id='mat-expansion-panel-header-3']");
        WebElement ActivedriverLabel = wait.until(ExpectedConditions.elementToBeClickable(ActiveDriverLocator));
        ActivedriverLabel.click();

        By excelBtnLocator_Driver = By.xpath("//img[@alt='Excel']");
        WebElement Vehicle_ExcelBtn_Driver = wait.until(ExpectedConditions.elementToBeClickable(excelBtnLocator_Driver));
        action.moveToElement(Vehicle_ExcelBtn_Driver).click().build().perform();
        WebElement Vehicle_CustomReport_Driver = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(text(),'Download Custom Report')]")));
        Vehicle_CustomReport_Driver.click();
        Thread.sleep(3000);
        String driverFileName = renameLatestFile("Active Driver");
        filterExcel(driverFileName, getDriverColumns());

        // Close current window and switch back
        driver.close();
        driver.switchTo().window(mainWindow);
        System.out.println("🔄 Switched back to main window: " + driver.getTitle());

        // ===== GPS Data Extraction =====
        WebElement SettingsBtn_GPS = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='billing']")));
        action.moveToElement(SettingsBtn_GPS).perform();

        By manageVehicleLocator_GPS = By.xpath("//li[@class='DeviceManagement roleBasedAccess']");
        WebElement Manage_Vehicle_GPS = wait.until(ExpectedConditions.elementToBeClickable(manageVehicleLocator_GPS));
        Manage_Vehicle_GPS.click();

        WebElement gpsTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table")));
        List<WebElement> gpsRows = gpsTable.findElements(By.xpath(".//tr"));

        Workbook gpsWorkbook = new XSSFWorkbook();
        Sheet gpsSheet = gpsWorkbook.createSheet("GPS Data");

        CellStyle headerStyle = gpsWorkbook.createCellStyle();
        Font headerFont = gpsWorkbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        CellStyle cellStyle = gpsWorkbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        String todayStr = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
        Date todayDate = new SimpleDateFormat("dd-MM-yyyy").parse(todayStr);

        List<String> headers = gpsRows.get(0)
                .findElements(By.xpath(".//th"))
                .stream()
                .map(WebElement::getText)
                .filter(h -> !h.equalsIgnoreCase("Action"))
                .collect(Collectors.toList());

        int rowIndex = 0;
        Row headerRow = gpsSheet.createRow(rowIndex++);
        int headerColIndex = 0;
        for (String header : headers) {
            Cell cell = headerRow.createCell(headerColIndex++);
            cell.setCellValue(header);
            cell.setCellStyle(headerStyle);
        }

        for (int i = 1; i < gpsRows.size(); i++) {
            List<WebElement> cells = gpsRows.get(i).findElements(By.xpath(".//td"));
            if (cells.size() > 0) {
                int lastContactIndex = headers.size() == cells.size() - 1 ? headers.size() - 1 : headers.size() - 2;
                String lastContactText = cells.get(lastContactIndex).getText();

                Date lastContactDate = null;
                try {
                    lastContactDate = new SimpleDateFormat("dd-MM-yyyy").parse(lastContactText.split(" ")[0]);
                } catch (Exception ignored) {}

                if (lastContactDate != null && lastContactDate.before(todayDate)) {
                    Row row = gpsSheet.createRow(rowIndex++);
                    int cellIndex = 0;
                    for (int j = 0; j < cells.size(); j++) {
                        String header = gpsRows.get(0).findElements(By.xpath(".//th")).get(j).getText();
                        if (!header.equalsIgnoreCase("Action")) {
                            Cell cell = row.createCell(cellIndex++);
                            cell.setCellValue(cells.get(j).getText());
                            cell.setCellStyle(cellStyle);
                        }
                    }
                }
            }
        }

        for (int i = 0; i < headers.size(); i++) {
            gpsSheet.autoSizeColumn(i);
        }

        String gpsFilePath = downloadPath + "\\GPS Data " + todayStr + ".xlsx";
        try (FileOutputStream fos = new FileOutputStream(gpsFilePath)) {
            gpsWorkbook.write(fos);
        }
        gpsWorkbook.close();
        System.out.println("✅ GPS Data saved with formatting: " + gpsFilePath);

        driver.quit();

        // ====== EMAIL Attachment Sending =======
        List<String> attachments = Arrays.asList(
            vehicleFileName,
            driverFileName,
            gpsFilePath
        );

        String fromEmail = "anshulgaur66@gmail.com";
        String toEmail = "srandhir770@gmail.com";
        String subject = "Cab Data Reports";
        String body = "Please find the attached cab data reports for today.";

        // NOTE: Use Gmail App Password, not your account password!
        String appPassword = "diqi sxyx zqjr nxhk"; // <-- replace this

        sendEmailWithAttachments(fromEmail, appPassword, toEmail, subject, body, attachments);

        String modeMessage = HEADLESS_MODE ? "headless mode" : "GUI mode";
        System.out.println("🚀 Script execution completed successfully in " + modeMessage + ".");
        System.exit(0);
    }

    // --- Utility methods ---
    private static String renameLatestFile(String prefix) throws Exception {
        String downloadPath = System.getProperty("user.home") + "\\Downloads";
        File downloadFolder = new File(downloadPath);

        File latestFile = Arrays.stream(downloadFolder.listFiles((dir, name) -> name.endsWith(".xlsx")))
                .max(Comparator.comparingLong(File::lastModified))
                .orElseThrow(() -> new RuntimeException("❌ No Excel file found in Downloads"));

        String today = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
        File renamedFile = new File(downloadPath + "\\" + prefix + " " + today + ".xlsx");

        Files.move(latestFile.toPath(), renamedFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        System.out.println("✅ File saved as: " + renamedFile.getName());
        return renamedFile.getAbsolutePath();
    }

    private static Set<String> getVehicleColumns() {
        return new HashSet<>(Arrays.asList(
                "Registration Number",
                "Overall Compliance Status",
                "Insurance Expiry Date",
                "Pollution Certificate Expiry Date",
                "Permit Expiry Date",
                "Road Tax Expiry Date",
                "Fitness Expiry Date",
                "EHS"
        ));
    }
    private static Set<String> getDriverColumns() {
        return new HashSet<>(Arrays.asList(
                "Name",
                "Overall Compliance Status",
                "Driver License Number",
                "Driver license expiry date",
                "Bgv Expiry Date",
                "Medical Verification Expiry Date",
                "Phone Numbers"
        ));
    }
    private static void filterExcel(String filePath, Set<String> requiredColumns) throws Exception {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            for (int colIndex = headerRow.getLastCellNum() - 1; colIndex >= 0; colIndex--) {
                Cell cell = headerRow.getCell(colIndex);
                if (cell != null) {
                    String colName = cell.getStringCellValue().trim();
                    if (!requiredColumns.contains(colName)) {
                        removeColumn(sheet, colIndex);
                    }
                }
            }
            int complianceColIndex = -1;
            for (Cell cell : headerRow) {
                if ("Overall Compliance Status".equalsIgnoreCase(cell.getStringCellValue().trim())) {
                    complianceColIndex = cell.getColumnIndex();
                    break;
                }
            }
            if (complianceColIndex != -1) {
                for (int rowIndex = sheet.getLastRowNum(); rowIndex > 0; rowIndex--) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        Cell statusCell = row.getCell(complianceColIndex);
                        if (statusCell != null && "Compliant".equalsIgnoreCase(statusCell.getStringCellValue().trim())) {
                            sheet.removeRow(row);
                            sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1);
                        }
                    }
                }
            }
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
        System.out.println("✅ Excel filtered: " + filePath);
    }
    private static void removeColumn(Sheet sheet, int colIndex) {
        for (Row row : sheet) {
            if (row.getCell(colIndex) != null) {
                row.removeCell(row.getCell(colIndex));
            }
            for (int i = colIndex; i < row.getLastCellNum(); i++) {
                Cell oldCell = row.getCell(i + 1);
                Cell newCell = row.getCell(i);
                if (oldCell != null) {
                    if (newCell == null) newCell = row.createCell(i);
                    cloneCell(oldCell, newCell);
                } else if (newCell != null) {
                    row.removeCell(newCell);
                }
            }
        }
    }
    private static void cloneCell(Cell oldCell, Cell newCell) {
        newCell.setCellStyle(oldCell.getCellStyle());
        switch (oldCell.getCellType()) {
            case STRING: newCell.setCellValue(oldCell.getStringCellValue()); break;
            case NUMERIC: newCell.setCellValue(oldCell.getNumericCellValue()); break;
            case BOOLEAN: newCell.setCellValue(oldCell.getBooleanCellValue()); break;
            case FORMULA: newCell.setCellFormula(oldCell.getCellFormula()); break;
            default: newCell.setBlank();
        }
    }

    // --- EMAIL SENDING UTILITY ----
    public static void sendEmailWithAttachments(String fromEmail, String password,
                                                String toEmail, String subject, String body,
                                                List<String> attachmentPaths) throws MessagingException {

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        Session session = Session.getInstance(props, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(fromEmail, password);
            }
        });
        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(fromEmail));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(toEmail));
        message.setSubject(subject);

        Multipart multipart = new MimeMultipart();

        // Body part
        MimeBodyPart messageBodyPart = new MimeBodyPart();
        messageBodyPart.setText(body);
        multipart.addBodyPart(messageBodyPart);

        // Attachments
        for (String filePath : attachmentPaths) {
            MimeBodyPart attachPart = new MimeBodyPart();
            File file = new File(filePath);
            DataSource source = new FileDataSource(file);
            attachPart.setDataHandler(new DataHandler(source));
            attachPart.setFileName(file.getName());
            multipart.addBodyPart(attachPart);
        }
        message.setContent(multipart);

        Transport.send(message);
        System.out.println("✅ Email sent successfully to " + toEmail);
    }
}
