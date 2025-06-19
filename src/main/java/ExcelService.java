


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Service class for Excel operations
 */
public class ExcelService {
    private static final Logger logger = LoggerFactory.getLogger(ExcelService.class);
    
    /**
     * Read student details from Excel file
     * @param filePath Path to the Excel file
     * @param sheetName Name of the sheet to read from (optional, uses first sheet if null)
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
                if (row == null) continue;
                
                Cell idCell = row.getCell(0);
                Cell nameCell = row.getCell(1);
                
                if (idCell != null && nameCell != null) {
                    String id = getCellValueAsString(idCell);
                    String name = getCellValueAsString(nameCell);
                    
                    if (!id.trim().isEmpty() && !name.trim().isEmpty()) {
                        students.add(new Student(id.trim(), name.trim()));
                    }
                }
            }
        }
        
        logger.info("Read {} students from Excel file", students.size());
        return students;
    }
    
    /**
     * Write teams to separate sheets in Excel file
     * @param teams List of student teams
     * @param outputFilePath Path for output Excel file
     * @throws IOException If file cannot be written
     */
    public void writeTeamsToExcel(List<List<Student>> teams, String outputFilePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            
            // Create cell styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle dataStyle = createDataStyle(workbook);
            
            for (int teamIndex = 0; teamIndex < teams.size(); teamIndex++) {
                List<Student> team = teams.get(teamIndex);
                String sheetName = "Team_" + (teamIndex + 1);
                
                Sheet sheet = workbook.createSheet(sheetName);
                
                // Create header row
                Row headerRow = sheet.createRow(0);
                Cell headerCell1 = headerRow.createCell(0);
                headerCell1.setCellValue("Student ID");
                headerCell1.setCellStyle(headerStyle);
                
                Cell headerCell2 = headerRow.createCell(1);
                headerCell2.setCellValue("Student Name");
                headerCell2.setCellStyle(headerStyle);
                
                // Add team data
                for (int studentIndex = 0; studentIndex < team.size(); studentIndex++) {
                    Student student = team.get(studentIndex);
                    Row dataRow = sheet.createRow(studentIndex + 1);
                    
                    Cell idCell = dataRow.createCell(0);
                    idCell.setCellValue(student.getId());
                    idCell.setCellStyle(dataStyle);
                    
                    Cell nameCell = dataRow.createCell(1);
                    nameCell.setCellValue(student.getName());
                    nameCell.setCellStyle(dataStyle);
                }
                
                // Auto-size columns
                sheet.autoSizeColumn(0);
                sheet.autoSizeColumn(1);
                
                logger.info("Created sheet '{}' with {} students", sheetName, team.size());
            }
            
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
                    // Handle numeric values (convert to string without decimal if it's a whole number)
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