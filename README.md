# Student Team Generator

A Maven-based Java application that reads student details from an Excel file and automatically splits them into teams of specified size, writing each team to separate sheets in an output Excel file.

### Note:

The application can also be executed as a standard Java application by running the ABC class. Before doing so, please perform a Maven Clean and Build to ensure all dependencies are correctly resolved. Until Maven-related concepts are fully integrated or covered, running the application in this plain Java mode is a suitable alternative for development and testing purposes.

## Features

- **Excel Integration**: Read student data from Excel files (.xlsx format)
- **Flexible Team Sizes**: Configurable team size (default: 10 students per team)
- **Multiple Output Sheets**: Each team is written to a separate sheet in the output Excel file
- **Professional Formatting**: Clean, formatted Excel output with headers and borders
- **Interactive Mode**: User-friendly command-line interface
- **Batch Processing**: Support for command-line arguments
- **Logging**: Comprehensive logging with SLF4J
- **Error Handling**: Robust error handling for file operations

## Project Structure

```
src/main/java/com/example/
├── StudentTeamGenerator.java          # Main application class
├── model/
│   └── Student.java                   # Student data model
├── service/
│   ├── ExcelService.java             # Excel read/write operations
│   └── TeamService.java              # Team management logic
└── util/
    └── SampleDataGenerator.java      # Generate sample test data
```

## Prerequisites

- Java 11 or higher
- Maven 3.6 or higher

## Dependencies

- **Apache POI**: Excel file processing
- **SLF4J**: Logging framework
- **JUnit**: Unit testing

## Getting Started

### 1. Clone and Build

```bash
# Clone the project (or create manually with the provided files)
cd student-team-generator

# Compile the project
mvn clean compile

# Run tests (optional)
mvn test

# Package the application
mvn clean package
```

### 2. Generate Sample Data (Optional)

To create a sample Excel file with student data for testing:

```bash
mvn exec:java -Dexec.mainClass="com.example.util.SampleDataGenerator"
```

This creates `sample_students.xlsx` with 35 sample students.

### 3. Run the Application

#### Interactive Mode

```bash
mvn exec:java -Dexec.mainClass="com.example.StudentTeamGenerator"
```

Follow the prompts to enter:

- Input Excel file path
- Output Excel file path
- Sheet name (optional - uses first sheet if not specified)
- Team size (default: 10)

#### Command Line Mode

```bash
mvn exec:java -Dexec.mainClass="com.example.StudentTeamGenerator" \
  -Dexec.args="input.xlsx output.xlsx [sheetName] [teamSize]"
```

**Examples:**

```bash
# Basic usage with defaults
mvn exec:java -Dexec.mainClass="com.example.StudentTeamGenerator" \
  -Dexec.args="students.xlsx student_teams.xlsx"

# Specify sheet name and team size
mvn exec:java -Dexec.mainClass="com.example.StudentTeamGenerator" \
  -Dexec.args="students.xlsx student_teams.xlsx Students 8"
```

## Input Excel Format

Your input Excel file should have the following format:

| Student ID | Student Name  |
| ---------- | ------------- |
| STU001     | Alice Johnson |
| STU002     | Bob Smith     |
| STU003     | Charlie Brown |
| ...        | ...           |

**Requirements:**

- First row should contain headers
- Column A: Student ID
- Column B: Student Name
- Save as .xlsx format

## Output Format

The application creates an output Excel file with multiple sheets:

- **Team_1**: First team (up to specified team size)
- **Team_2**: Second team (up to specified team size)
- **Team_N**: Additional teams as needed

Each sheet contains:

- Formatted headers with blue background
- Student ID and Name columns
- Bordered cells for professional appearance
- Auto-sized columns

## Example Usage

1. **Create sample data:**

   ```bash
   mvn exec:java -Dexec.mainClass="com.example.util.SampleDataGenerator"
   ```

2. **Generate teams:**

   ```bash
   mvn exec:java -Dexec.mainClass="com.example.StudentTeamGenerator" \
     -Dexec.args="sample_students.xlsx teams_output.xlsx"
   ```

3. **Result:**
   - Input: 35 students in `sample_students.xlsx`
   - Output: 4 teams in `teams_output.xlsx` (3 teams of 10, 1 team of 5)

## Customization Options

### Change Team Size

Modify the team size by passing it as the 4th argument
