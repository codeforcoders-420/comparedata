package splittherates.ratespan;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Hello world!
 */
public class RateHistoryConverter {
	public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\rajas\\Desktop\\ratecompare\\Rate History.xlsx";
        String outputFilePath = "C:\\Users\\rajas\\Desktop\\ratecompare\\Output\\Output Rate History.xlsx";

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet inputSheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet
            List<RateSpan> rateSpans = processRateHistory(inputSheet);

            writeOutputFile(outputFilePath, rateSpans);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<RateSpan> processRateHistory(Sheet sheet) {
        List<RateSpan> rateSpans = new ArrayList<>();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy");

        Row headerRow = sheet.getRow(0); // Header row (containing month/year)
        int columnStartIndex = 3; // Column D is index 3

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            String code = row.getCell(0).getStringCellValue();
            Double currentRate = null;
            LocalDate spanStartDate = null;

            for (int colIndex = columnStartIndex; colIndex < row.getLastCellNum(); colIndex++) {
                LocalDate currentDate = parseDateFromHeader(headerRow.getCell(colIndex));
                Double rate = row.getCell(colIndex).getNumericCellValue();

                if (currentRate == null || !currentRate.equals(rate)) {
                    if (currentRate != null) {
                        LocalDate endDate = currentDate.minusDays(1);
                        rateSpans.add(new RateSpan(code, spanStartDate, endDate, currentRate));
                    }
                    currentRate = rate;
                    spanStartDate = currentDate;
                }

                // Special case for the last column
                if (colIndex == row.getLastCellNum() - 1) {
                    LocalDate endDate = LocalDate.of(9999, 12, 31); // Open-ended date
                    rateSpans.add(new RateSpan(code, spanStartDate, endDate, currentRate));
                }
            }
        }

        return rateSpans;
    }

    private static LocalDate parseDateFromHeader(Cell cell) {
        if (cell == null) {
            throw new IllegalArgumentException("Cell is null.");
        }

        // Check if the cell contains a numeric or string value
        if (cell.getCellType() == CellType.STRING) {
            // Handle String cell value (MM/yyyy format)
            String headerDateStr = cell.getStringCellValue();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/yyyy");
            return LocalDate.parse("01/" + headerDateStr, formatter); // Parsing with a fixed first day of the month
        } else if (cell.getCellType() == CellType.NUMERIC) {
            // Handle Numeric cell value (which could be a date)
            if (DateUtil.isCellDateFormatted(cell)) {
                // If it's a date, retrieve it as a LocalDate
                Date date = cell.getDateCellValue();
                return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            } else {
                throw new IllegalStateException("Unexpected numeric cell that is not formatted as a date.");
            }
        } else {
            throw new IllegalStateException("Unexpected cell type: " + cell.getCellType());
        }
    }


    private static void writeOutputFile(String outputFilePath, List<RateSpan> rateSpans) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputFilePath)) {

            Sheet outputSheet = workbook.createSheet("Rate History");
            createHeader(outputSheet);

            int rowNum = 1;
            for (RateSpan span : rateSpans) {
                Row row = outputSheet.createRow(rowNum++);
                row.createCell(0).setCellValue(span.getCode());
                row.createCell(1).setCellValue(span.getStartDate().format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
                row.createCell(2).setCellValue(span.getEndDate().format(DateTimeFormatter.ofPattern("MM/dd/yyyy")));
                row.createCell(3).setCellValue(span.getRate());
            }

            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void createHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Code");
        headerRow.createCell(1).setCellValue("Start Date");
        headerRow.createCell(2).setCellValue("End Date");
        headerRow.createCell(3).setCellValue("Rate");
    }

    static class RateSpan {
        private final String code;
        private final LocalDate startDate;
        private final LocalDate endDate;
        private final Double rate;

        public RateSpan(String code, LocalDate startDate, LocalDate endDate, Double rate) {
            this.code = code;
            this.startDate = startDate;
            this.endDate = endDate;
            this.rate = rate;
        }

        public String getCode() {
            return code;
        }

        public LocalDate getStartDate() {
            return startDate;
        }

        public LocalDate getEndDate() {
            return endDate;
        }

        public Double getRate() {
            return rate;
        }
    }
}
