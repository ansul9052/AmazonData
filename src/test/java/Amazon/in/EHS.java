package Amazon.in;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class EHS {

    private static final Pattern VEHICLE_PATTERN = Pattern.compile("([A-Z]{2})(\\d{2})([A-Z]{1,2})(\\d{4})");
    private static final DateTimeFormatter SITE_FORMATTER = DateTimeFormatter.ofPattern("MMM d, yyyy, h:mm:ss a", Locale.ENGLISH);

    static class VehicleInfo {
        String vehicleNo;
        String driverName;
        String ehsDateTime;

        VehicleInfo(String vehicleNo, String driverName) {
            this.vehicleNo = vehicleNo;
            this.driverName = driverName;
        }

        VehicleInfo(String vehicleNo, String driverName, String ehsDateTime) {
            this.vehicleNo = vehicleNo;
            this.driverName = driverName;
            this.ehsDateTime = ehsDateTime;
        }
    }

    public static void main(String[] args) throws Exception {

        boolean runHeadless = true;
        ChromeOptions options = new ChromeOptions();
        if (runHeadless) {
            options.addArguments("--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu", "--window-size=1920,1080");
        }
        options.addArguments("--disable-blink-features=AutomationControlled",
                "--disable-extensions", "--disable-plugins", "--disable-images", "--disable-background-timer-throttling");

        WebDriver driver = new ChromeDriver(options);
        if (runHeadless) driver.manage().window().setSize(new Dimension(1920, 1080));

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));

        String downloadsPath = Paths.get(System.getProperty("user.home"), "Downloads").toString();
        String todayStr = new SimpleDateFormat("dd-MM-yyyy").format(new Date());
        String fileName = "EHS " + todayStr + ".xlsx";
        String filePath = Paths.get(downloadsPath, fileName).toString();

        try {
            // Login & Navigate
            driver.get("https://amz.moveinsync.com/Hyd/");
            driver.findElement(By.xpath("//input[@class='loginInput']")).sendKeys("euro-ashraf-ops@dummy.com");
            driver.findElement(By.xpath("//input[@class='loginInput passInput']")).sendKeys("Euro@789456");
            driver.findElement(By.xpath("//input[@class='loginBtn loginEnabled']")).click();

            WebElement refreshBtn = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@title='Click to refresh now']")));
            refreshBtn.click();

            WebElement settingsBtn = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("billing")));
            new Actions(driver).moveToElement(settingsBtn).perform();

            WebElement manageVehicle = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li[@class='CabManagement roleBasedAccess']")));
            manageVehicle.click();

            wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt(By.id("setting-dashboards")));
            WebElement table = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//table")));

            List<VehicleInfo> vehicleList = extractVehicleData(table);
            writeInitialExcel(filePath, vehicleList);
            System.out.println("✅ Initial Excel file created: " + filePath);

            if (runHeadless) Thread.sleep(1000);

            WebElement manageCompliance = wait.until(ExpectedConditions.visibilityOfElementLocated(
                    By.cssSelector(".manage-compliance-link.ng-tns-c2299869914-0")));
            String mainWindow = driver.getWindowHandle();
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", manageCompliance);

            wait.until(ExpectedConditions.numberOfWindowsToBe(2));
            for (String window : driver.getWindowHandles()) {
                if (!window.equals(mainWindow)) {
                    driver.switchTo().window(window);
                    break;
                }
            }

            WebElement vehicleAudit = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@class='header-links']")));
            vehicleAudit.click();

            WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#mat-input-1")));
            searchBox.sendKeys("EHS");

            WebElement vehicleSubmit = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("#mat-input-0")));
            List<VehicleInfo> validVehicles = processVehiclesOptimized(vehicleList, vehicleSubmit, wait);

            LocalDate today = LocalDate.now();
            LocalDate yesterday = today.minusDays(1);
            validVehicles.sort((v1, v2) -> {
                LocalDate date1 = LocalDate.MIN, date2 = LocalDate.MIN;
                try { date1 = LocalDate.parse(v1.ehsDateTime, SITE_FORMATTER); } catch (Exception ignored) {}
                try { date2 = LocalDate.parse(v2.ehsDateTime, SITE_FORMATTER); } catch (Exception ignored) {}
                if (date1.equals(today) && date2.equals(yesterday)) return -1;
                if (date1.equals(yesterday) && date2.equals(today)) return 1;
                return date2.compareTo(date1);
            });

            writeFinalExcel(filePath, validVehicles);
            System.out.println("✅ Final Excel saved: " + filePath);

            // Send Email after Excel creation
            sendEmailWithAttachments(
                    "anshulgaur66@gmail.com",            // From email
                    "diqi sxyx zqjr nxhk",           // 🔹 Replace with Gmail App Password
                    "srandhir770@gmail.com",              // To email
                    "EHS Report - " + todayStr,           // Subject
                    "Please find attached EHS report for " + todayStr,
                    Collections.singletonList(filePath)   // Attachment
            );

        } finally {
            driver.quit();
        }
    }

    // ===== Extract Data =====
    private static List<VehicleInfo> extractVehicleData(WebElement table) {
        List<WebElement> rows = table.findElements(By.xpath(".//tr[position()>1]"));
        List<VehicleInfo> vehicleList = new ArrayList<>();
        for (WebElement row : rows) {
            List<WebElement> cols = row.findElements(By.tagName("td"));
            if (cols.size() >= 7) {
                String vehicleNo = formatVehicleNumber(cols.get(2).getText().trim());
                String driverName = cols.get(6).getText().trim();
                vehicleList.add(new VehicleInfo(vehicleNo, driverName));
            }
        }
        return vehicleList;
    }

    // ===== Process Vehicles =====
    private static List<VehicleInfo> processVehiclesOptimized(List<VehicleInfo> vehicleList,
                                                              WebElement vehicleSubmit, WebDriverWait wait) {
        List<VehicleInfo> validVehicles = new ArrayList<>();
        LocalDate today = LocalDate.now(), yesterday = today.minusDays(1);
        int processed = 0, total = vehicleList.size();

        for (VehicleInfo vehicle : vehicleList) {
            try {
                vehicleSubmit.clear();
                vehicleSubmit.sendKeys(vehicle.vehicleNo);
                vehicleSubmit.sendKeys(Keys.ENTER);

                WebElement changeTimeCell = wait.until(ExpectedConditions.visibilityOfElementLocated(
                        By.xpath("(//td[contains(@class,'cdk-column-updatedTime')])[1]")));
                String changeTimeText = changeTimeCell.getText().trim();

                try {
                    LocalDate changeDate = LocalDate.parse(changeTimeText, SITE_FORMATTER);
                    if (changeDate.equals(today) || changeDate.equals(yesterday)) {
                        validVehicles.add(new VehicleInfo(vehicle.vehicleNo.replace("-", ""), vehicle.driverName, changeTimeText));
                    }
                } catch (Exception ignored) {}
            } catch (Exception ignored) {}

            processed++;
            System.out.printf("\rProcessed %d/%d", processed, total);
        }
        System.out.println();
        return validVehicles;
    }

    // ===== Excel Creation =====
    private static void writeInitialExcel(String filePath, List<VehicleInfo> vehicleList) throws Exception {
        writeExcel(filePath, vehicleList, false);
    }

    private static void writeFinalExcel(String filePath, List<VehicleInfo> validVehicles) throws Exception {
        writeExcel(filePath, validVehicles, true);
    }

    private static void writeExcel(String filePath, List<VehicleInfo> vehicles, boolean withDates) throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Vehicle Data");
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle cellStyle = createCellStyle(workbook);

            Row headerRow = sheet.createRow(0);
            String[] headers = {"Vehicle No", "Driver", "EHS Date & Time"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            int rowNum = 1;
            for (VehicleInfo v : vehicles) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(v.vehicleNo);
                row.createCell(1).setCellValue(v.driverName);
                row.createCell(2).setCellValue(withDates && v.ehsDateTime != null ? v.ehsDateTime : "");
                for (int i = 0; i <= 2; i++) row.getCell(i).setCellStyle(cellStyle);
            }
            for (int i = 0; i < 3; i++) sheet.autoSizeColumn(i);

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    // ===== Styles =====
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    // ===== Vehicle Number Formatter =====
    public static String formatVehicleNumber(String raw) {
        if (raw == null || raw.isEmpty()) return raw;
        raw = raw.replaceAll("[^A-Za-z0-9]", "").toUpperCase();
        Matcher matcher = VEHICLE_PATTERN.matcher(raw);
        if (matcher.matches()) {
            return matcher.group(1) + "-" + matcher.group(2) + "-" + matcher.group(3) + "-" + matcher.group(4);
        }
        return raw;
    }

    // ===== Email Utility =====
    public static void sendEmailWithAttachments(String fromEmail, String password,
                                                String toEmail, String subject, String body,
                                                List<String> attachmentPaths) throws MessagingException {

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");

        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(fromEmail, password);
            }
        });

        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(fromEmail));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(toEmail));
        message.setSubject(subject);

        Multipart multipart = new MimeMultipart();
        MimeBodyPart textPart = new MimeBodyPart();
        textPart.setText(body);
        multipart.addBodyPart(textPart);

        for (String filePath : attachmentPaths) {
            MimeBodyPart attachment = new MimeBodyPart();
            DataSource source = new FileDataSource(new File(filePath));
            attachment.setDataHandler(new DataHandler(source));
            attachment.setFileName(new File(filePath).getName());
            multipart.addBodyPart(attachment);
        }

        message.setContent(multipart);
        Transport.send(message);
        System.out.println("✅ Email sent successfully to " + toEmail);
    }
}
