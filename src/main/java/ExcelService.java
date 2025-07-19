
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Service class for Excel operations
 */
public class ExcelService {
    private static final Logger logger = LoggerFactory.getLogger(ExcelService.class);

    /**
     * Read student details from Excel file
     * 
     * @param filePath  Path to the Excel file
     * @param sheetName Name of the sheet to read from (optional, uses first sheet
     *                  if null)
     * @return List of students
     * @throws IOException If file cannot be read
     */
    public List<Student> readStudentsFromExcel(String filePath, String sheetName) throws IOException {
        List<Student> students = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
                Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = (sheetName != null) ? workbook.getSheet(sheetName) : workbook.getSheetAt(0);

            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            logger.info("Reading students from sheet: {}", sheet.getSheetName());

            // Skip header row (assuming first row contains headers)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null)
                    continue;

                Cell idCell = row.getCell(0);
                Cell regNoCell = row.getCell(1);
                Cell nameCell = row.getCell(2);
                Cell branchCell = row.getCell(3);

                if (idCell != null && nameCell != null) {
                    String id = getCellValueAsString(idCell);
                    String name = getCellValueAsString(nameCell);
                    String regNo = getCellValueAsString(regNoCell);
                    String branch = getCellValueAsString(branchCell);

                    if (!id.trim().isEmpty() && !name.trim().isEmpty()) {
                        students.add(new Student(id.trim(), regNo.trim(), name.trim(), branch.trim()));
                    }
                }
            }
        }

        logger.info("Read {} students from Excel file", students.size());
        return students;
    }

    /**
     * Write teams to separate sheets in Excel file
     * 
     * @param teams          List of student teams
     * @param outputFilePath Path for output Excel file
     * @throws IOException If file cannot be written
     */
    public void writeTeamsToExcel(List<List<Student>> teams, String outputFilePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            List<Student> allStudents = teams.stream()
                    .flatMap(List::stream)
                    .collect(Collectors.toList());
            // Create cell styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dataStyle = createDataStyle(workbook);
            String sheetName = "Student-Teams";
            Sheet sheet = workbook.createSheet(sheetName);
            for (int i = 0; i <= allStudents.size(); i++) {

                // Create header row
                if (i == 0) {
                    Row headerRow = sheet.createRow(0);
                    Cell headerCell1 = headerRow.createCell(0);
                    headerCell1.setCellValue("Student ID");
                    headerCell1.setCellStyle(headerStyle);

                    Cell headerCell2 = headerRow.createCell(1);
                    headerCell2.setCellValue("Student Name");
                    headerCell2.setCellStyle(headerStyle);

                    Cell headerCell3 = headerRow.createCell(2);
                    headerCell3.setCellValue("RegId");
                    headerCell3.setCellStyle(headerStyle);
                    Cell headerCell4 = headerRow.createCell(3);
                    headerCell4.setCellValue("Dept");
                    headerCell4.setCellStyle(headerStyle);

                    Cell headerCell5 = headerRow.createCell(4);
                    headerCell5.setCellValue("Team Name");
                    headerCell5.setCellStyle(headerStyle);
                } else {

                    Student student = allStudents.get(i - 1);
                    Row dataRow = sheet.createRow(i);

                    Cell idCell = dataRow.createCell(0);
                    idCell.setCellValue(student.getId());
                    idCell.setCellStyle(dataStyle);

                    Cell nameCell = dataRow.createCell(1);
                    nameCell.setCellValue(student.getName());
                    nameCell.setCellStyle(dataStyle);

                    Cell regIdCell = dataRow.createCell(2);
                    regIdCell.setCellValue(student.getRegId());
                    regIdCell.setCellStyle(dataStyle);

                    Cell deptCell = dataRow.createCell(3);
                    deptCell.setCellValue(student.getDept());
                    deptCell.setCellStyle(dataStyle);

                    Cell teamCell = dataRow.createCell(4);
                    teamCell.setCellValue(student.getTeam());
                    teamCell.setCellStyle(dataStyle);
                }

            }

            // Auto-size columns
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.autoSizeColumn(4);

            // Write to file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        }

        logger.info("Successfully wrote {} teams to Excel file: {}", teams.size(), outputFilePath);
    }

    /**
     * Create header cell style
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Create data cell style
     */
    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Convert cell value to string
     */
    private String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // Handle numeric values (convert to string without decimal if it's a whole
                    // number)
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == Math.floor(numericValue)) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}