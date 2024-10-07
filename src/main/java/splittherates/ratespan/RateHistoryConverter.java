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

            // Output logic here (writing to another file or processing the result)
            // For now, just printing the spans
            for (Map.Entry<String, List<RateSpan>> entry : rateSpansByCode.entrySet()) {
                System.out.println("Code: " + entry.getKey());
                for (RateSpan span : entry.getValue()) {
                    System.out.println(span);
                }
            }

            // Optionally, write the output to an Excel file (not covered here)
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
}
