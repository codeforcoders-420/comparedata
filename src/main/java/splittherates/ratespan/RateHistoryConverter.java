package splittherates.ratespan;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class RateHistoryConverter {

    private static final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

    public static void main(String[] args) throws IOException {
        FileInputStream inputStream = new FileInputStream("/mnt/data/Rate History.xlsx"); // Adjust path accordingly
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Output workbook and sheet
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Rate History Output");

        Map<String, List<RateSpan>> rateSpansMap = new LinkedHashMap<>();

        // Read header row (for the dates)
        Row headerRow = sheet.getRow(0);
        int numColumns = headerRow.getPhysicalNumberOfCells();

        // Read data starting from the second row
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            String proc = row.getCell(0).getStringCellValue(); // PROC
            String mod = row.getCell(1) != null ? row.getCell(1).getStringCellValue() : ""; // MOD
            String mod2 = row.getCell(2) != null ? row.getCell(2).getStringCellValue() : ""; // MOD2

            String compositeKey = proc + "_" + mod + "_" + mod2; // Composite key

            List<RateSpan> rateSpans = new ArrayList<>();

            for (int colIndex = 3; colIndex < numColumns; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) continue;

                // Get the start date from the header
                Cell headerCell = headerRow.getCell(colIndex);
                LocalDate startDate = parseDateFromHeader(headerCell);

                // Get the end date (minus one day from the next start date or 12/31/9999 for the last column)
                LocalDate endDate;
                if (colIndex + 1 < numColumns) {
                    Cell nextHeaderCell = headerRow.getCell(colIndex + 1);
                    endDate = parseDateFromHeader(nextHeaderCell).minusDays(1);
                } else {
                    endDate = LocalDate.of(9999, 12, 31); // Open-ended date for the last column
                }

                // Get the rate from the current cell
                double rate = cell.getNumericCellValue();

                // Add to rate spans list
                rateSpans.add(new RateSpan(startDate, endDate, rate));
            }

            // Combine spans where rates are the same
            rateSpansMap.put(compositeKey, combineRateSpans(rateSpans));
        }

        // Write the output file
        int outputRowNum = 0;
        for (Map.Entry<String, List<RateSpan>> entry : rateSpansMap.entrySet()) {
            String[] keyParts = entry.getKey().split("_");
            String proc = keyParts[0];
            String mod = keyParts[1];
            String mod2 = keyParts[2];

            for (RateSpan span : entry.getValue()) {
                Row outputRow = outputSheet.createRow(outputRowNum++);
                outputRow.createCell(0).setCellValue(proc); // PROC
                outputRow.createCell(1).setCellValue(mod);  // MOD
                outputRow.createCell(2).setCellValue(mod2); // MOD2
                outputRow.createCell(3).setCellValue(span.startDate.format(formatter)); // Start Date
                outputRow.createCell(4).setCellValue(span.endDate.format(formatter)); // End Date
                outputRow.createCell(5).setCellValue(span.rate); // Rate
            }
        }

        FileOutputStream outputStream = new FileOutputStream("/mnt/data/Output Rate History.xlsx"); // Adjust path
        outputWorkbook.write(outputStream);
        outputStream.close();

        workbook.close();
        outputWorkbook.close();
    }

    // Method to combine rate spans where rates are consecutive and equal
    private static List<RateSpan> combineRateSpans(List<RateSpan> rateSpans) {
        List<RateSpan> combinedSpans = new ArrayList<>();
        if (rateSpans.isEmpty()) return combinedSpans;

        RateSpan currentSpan = rateSpans.get(0);
        for (int i = 1; i < rateSpans.size(); i++) {
            RateSpan nextSpan = rateSpans.get(i);
            if (currentSpan.rate == nextSpan.rate) {
                currentSpan.endDate = nextSpan.endDate; // Extend the current span
            } else {
                combinedSpans.add(currentSpan);
                currentSpan = nextSpan; // Start a new span
            }
        }
        combinedSpans.add(currentSpan); // Add the last span

        return combinedSpans;
    }

    // Method to parse the date from the header (Month/Year format)
    private static LocalDate parseDateFromHeader(Cell cell) {
        if (cell.getCellType() == CellType.STRING) {
            String headerDateStr = cell.getStringCellValue();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/yyyy");
            return LocalDate.parse("01/" + headerDateStr, formatter); // Parse as first day of the month
        } else if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        }
        throw new IllegalStateException("Invalid date format in header cell");
    }

    // Class to hold rate spans
    static class RateSpan {
        LocalDate startDate;
        LocalDate endDate;
        double rate;

        RateSpan(LocalDate startDate, LocalDate endDate, double rate) {
            this.startDate = startDate;
            this.endDate = endDate;
            this.rate = rate;
        }
    }
}
