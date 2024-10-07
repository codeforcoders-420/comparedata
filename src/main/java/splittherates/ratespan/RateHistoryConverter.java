import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class RateHistoryConverter {

    private static final DateTimeFormatter headerFormatter = DateTimeFormatter.ofPattern("MMM yyyy");
    private static final DateTimeFormatter fullFormatter = DateTimeFormatter.ofPattern("MMMM yyyy");
    private static final DateTimeFormatter outputDateFormatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

    public static void main(String[] args) throws IOException {
        String inputFilePath = "/mnt/data/Rate History.xlsx";  // Update with your file path
        String outputFilePath = "/mnt/data/Output Rate History.xlsx";  // Update with your output file path

        // Load the input Excel file
        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Get the header row (assume the first row is the header)
            Row headerRow = sheet.getRow(0);
            List<LocalDate> dates = new ArrayList<>();

            // Parse the dates from the header starting at column D (index 3)
            for (int col = 3; col < headerRow.getLastCellNum(); col++) {
                Cell cell = headerRow.getCell(col);
                dates.add(parseDateFromHeader(cell));
            }

            // Map to hold spans by code
            Map<String, List<RateSpan>> rateSpansByCode = new HashMap<>();

            // Iterate over the rows (start from the second row for data)
            for (int rowIdx = 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) continue;  // Skip empty rows

                // Generate rate spans for this row
                List<RateSpan> rateSpans = generateRateSpansForRow(row, 3, dates);  // Start reading from column D (index 3)

                // Build a key based on PROC + MOD + MOD2
                String key = row.getCell(0).getStringCellValue().trim();  // PROC code
                if (row.getCell(1) != null) {
                    key += "-" + row.getCell(1).getStringCellValue().trim();  // MOD
                }
                if (row.getCell(2) != null) {
                    key += "-" + row.getCell(2).getStringCellValue().trim();  // MOD2
                }

                // Add the spans to the map
                rateSpansByCode.put(key, rateSpans);
            }

            // Write the rate spans to an Excel file
            writeRateSpansToExcel(rateSpansByCode, outputFilePath);
        }
    }

    // Method to parse the date from the header (handling both full and abbreviated month names)
    private static LocalDate parseDateFromHeader(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            String headerDateStr = cell.getStringCellValue().trim();  // Trim any extra spaces

            try {
                // First attempt: Parse as abbreviated month name (MMM yyyy)
                YearMonth yearMonth = YearMonth.parse(headerDateStr, headerFormatter);
                return yearMonth.atDay(1);  // Convert to LocalDate as the 1st of the month
            } catch (DateTimeParseException e1) {
                try {
                    // Second attempt: Parse as full month name (MMMM yyyy)
                    YearMonth yearMonth = YearMonth.parse(headerDateStr, fullFormatter);
                    return yearMonth.atDay(1);  // Convert to LocalDate as the 1st of the month
                } catch (DateTimeParseException e2) {
                    throw new IllegalStateException("Invalid date format in header: " + headerDateStr, e2);
                }
            }
        }
        throw new IllegalStateException("Invalid date format in header cell");
    }

    // Method to safely get the rate from the cell
    private static Double getRateFromCell(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) {
            return null;  // Handle blank cells
        }

        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();  // Return numeric value
        } else if (cell.getCellType() == CellType.STRING) {
            // Try to parse the string value as a number (if possible)
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return null;  // Return null or handle invalid numeric string
            }
        }
        return null;  // Default return for unsupported cell types
    }

    // Method to generate rate spans for a row, considering blank cells and handling the conditions
    private static List<RateSpan> generateRateSpansForRow(Row row, int startColumn, List<LocalDate> dates) {
        List<RateSpan> rateSpans = new ArrayList<>();
        String procCode = row.getCell(0).getStringCellValue().trim();
        String mod = row.getCell(1) != null ? row.getCell(1).getStringCellValue().trim() : "";
        String mod2 = row.getCell(2) != null ? row.getCell(2).getStringCellValue().trim() : "";

        Double lastRate = null;
        LocalDate spanStart = null;

        for (int col = startColumn; col < row.getLastCellNum(); col++) {
            Cell cell = row.getCell(col);
            Double currentRate = getRateFromCell(cell);
            LocalDate currentDate = dates.get(col - startColumn);

            // Skip blank rates at the start
            if (lastRate == null && currentRate == null) {
                continue;  // Skip until we find the first rate
            }

            // If we find a rate, start the span
            if (lastRate == null && currentRate != null) {
                lastRate = currentRate;
                spanStart = currentDate;
                continue;
            }

            // If the current rate is different or if we hit a blank after having a rate
            if (currentRate == null || !currentRate.equals(lastRate)) {
                // Terminate the previous span
                LocalDate spanEnd = (currentRate == null) ? currentDate.minusDays(1) : currentDate.minusDays(1);
                rateSpans.add(new RateSpan(procCode, mod, mod2, lastRate, spanStart, spanEnd));

                // Start a new span only if the current rate is non-null
                if (currentRate != null) {
                    lastRate = currentRate;
                    spanStart = currentDate;
                } else {
                    lastRate = null;  // Reset the lastRate if we hit a blank
                    spanStart = null;
                }
            }
        }

        // Handle the last span that should go until 12/31/9999
        if (lastRate != null) {
            rateSpans.add(new RateSpan(procCode, mod, mod2, lastRate, spanStart, LocalDate.of(9999, 12, 31)));
        }

        return rateSpans;
    }

    // New method to write rate spans to an Excel file
    private static void writeRateSpansToExcel(Map<String, List<RateSpan>> rateSpansByCode, String outputFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Rate History");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("PROC");
        headerRow.createCell(1).setCellValue("MOD");
        headerRow.createCell(2).setCellValue("MOD2");
        headerRow.createCell(3).setCellValue("Rate");
        headerRow.createCell(4).setCellValue("Start Date");
        headerRow.createCell(5).setCellValue("End Date");

        int rowIndex = 1;  // Start writing from the second row

        // Loop over the rate spans and write each to a new row
        for (Map.Entry<String, List<RateSpan>> entry : rateSpansByCode.entrySet()) {
            List<RateSpan> rateSpans = entry.getValue();

            for (RateSpan span : rateSpans) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(span.getProcCode());
                row.createCell(1).setCellValue(span.getMod());
                row.createCell(2).setCellValue(span.getMod2());
                row.createCell(3).setCellValue(span.getRate());
                row.createCell(4).setCellValue(span.getStartDate().format(outputDateFormatter));
                row.createCell(5).setCellValue(span.getEndDate().format(outputDateFormatter));
            }
        }

        // Auto-size columns for better readability
        for (int i = 0; i <= 5; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
            workbook.write(fos);
        }

        // Close the workbook
        workbook.close();

        System.out.println("Output written to: " + outputFilePath);
    }
}

// The RateSpan class definition
class RateSpan {
    private String procCode;
    private String mod;
    private String mod2;
    private Double rate;
    private LocalDate startDate;
    private LocalDate endDate;

    public RateSpan(String procCode, String mod, String mod2, Double rate, LocalDate startDate, LocalDate endDate) {
        this.procCode = procCode;
        this.mod = mod;
        this.mod2 = mod2;
        this.rate = rate;
        this.startDate = startDate;
        this.endDate = endDate;
    }

    // Getters
    public String getProcCode() {
        return procCode;
    }

    public String getMod() {
        return mod;
    }

    public String getMod2() {
        return mod2;
    }

    public Double getRate() {
        return rate;
    }

    public LocalDate getStartDate() {
        return startDate;
    }

    public LocalDate getEndDate() {
        return endDate;
    }

    @Override
    public String toString() {
        return String.format("RateSpan[PROC=%s, MOD=%s, MOD2=%s, Rate=%.2f, Start=%s, End=%s]",
                procCode, mod, mod2, rate, startDate, endDate);
    }
}
