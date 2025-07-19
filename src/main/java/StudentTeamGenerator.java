
import java.io.IOException;
import java.util.List;
import java.util.Scanner;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Main application class for Student Team Generator
 */
public class StudentTeamGenerator {
    private static final Logger logger = LoggerFactory.getLogger(StudentTeamGenerator.class);

    // Tight coupling: Directly using the concrete class ExcelService binds us to a
    // specific implementation.
    // This makes testing and future changes harder since we're not leveraging
    // abstraction.
    private ExcelService excelService;

    // Loose coupling: By referencing ITeamsService (an interface), we depend on an
    // abstraction.
    // This promotes flexibility, easier testing (e.g., mocking), and better
    // adherence to design principles like SOLID.
    private ITeamsService teamService;

    public StudentTeamGenerator() {
        this.excelService = new ExcelService();
        this.teamService = new RandomTeamsService();
    }

    public static void main(String[] args) {
        StudentTeamGenerator app = new StudentTeamGenerator();

        if (args.length >= 2) {
            // Command line arguments provided
            String inputFile = args[0];
            String outputFile = args[1];
            String sheetName = args.length > 2 ? args[2] : null;
            int teamSize = args.length > 3 ? Integer.parseInt(args[3]) : 10;

            app.processStudentTeams(inputFile, outputFile, sheetName, teamSize);
        } else {
            // Interactive mode
            app.runInteractiveMode();
        }
    }

    /**
     * Run the application in interactive mode
     */
    public void runInteractiveMode() {
        Scanner scanner = new Scanner(System.in);

        try {
            System.out.println("=== Student Team Generator ===");
            System.out.println();

            // Get input file path
            System.out.print("Enter input Excel file path: ");
            String inputFile = scanner.nextLine().trim();

            // Get output file path
            System.out.print("Enter output Excel file path: ");
            String outputFile = scanner.nextLine().trim();

            // Get sheet name (optional)
            System.out.print("Enter sheet name to read from (press Enter for first sheet): ");
            String sheetName = scanner.nextLine().trim();
            sheetName = sheetName.isEmpty() ? null : sheetName;

            // Get team size
            System.out.print("Enter team size (default: 10): ");
            String teamSizeInput = scanner.nextLine().trim();
            int teamSize = teamSizeInput.isEmpty() ? 10 : Integer.parseInt(teamSizeInput);

            System.out.println();
            processStudentTeams(inputFile, outputFile, sheetName, teamSize);

        } catch (Exception e) {
            logger.error("Error in interactive mode: {}", e.getMessage(), e);
            System.err.println("Error: " + e.getMessage());
        } finally {
            scanner.close();
        }
    }

    /**
     * Process student teams with given parameters
     */
    public void processStudentTeams(String inputFile, String outputFile, String sheetName, int teamSize) {
        try {
            logger.info("Starting student team generation process");
            logger.info("Input file: {}", inputFile);
            logger.info("Output file: {}", outputFile);
            logger.info("Sheet name: {}", sheetName != null ? sheetName : "First sheet");
            logger.info("Team size: {}", teamSize);

            // Step 1: Read students from Excel
            System.out.println("Reading students from Excel file...");
            List<Student> students = excelService.readStudentsFromExcel(inputFile, sheetName);

            if (students.isEmpty()) {
                System.out.println("No students found in the Excel file!");
                return;
            }

            System.out.printf("Found %d students in the Excel file%n", students.size());

            // Step 2: Split students into teams
            System.out.println("Creating teams...");
            List<List<Student>> teams = teamService.splitIntoTeams(students, teamSize);

            // // Step 3: Display team statistics
            // System.out.println();
            // System.out.println(teamService.getTeamStatistics(teams));
            // System.out.println();

            // Step 4: Write teams to Excel
            System.out.println("Writing teams to Excel file...");
            excelService.writeTeamsToExcel(teams, outputFile);

            System.out.println("✓ Successfully created team assignments!");
            System.out.printf("✓ Output file: %s%n", outputFile);
            System.out.printf("✓ Created %d teams%n", teams.size());

        } catch (IOException e) {
            logger.error("File operation error: {}", e.getMessage(), e);
            System.err.println("File error: " + e.getMessage());
        } catch (Exception e) {
            logger.error("Unexpected error: {}", e.getMessage(), e);
            System.err.println("Error: " + e.getMessage());
        }
    }

    /**
     * Process with default parameters (for testing)
     */
    public void processWithDefaults() {
        String inputFile = "students.xlsx";
        String outputFile = "student_teams.xlsx";
        processStudentTeams(inputFile, outputFile, null, 10);
    }
}