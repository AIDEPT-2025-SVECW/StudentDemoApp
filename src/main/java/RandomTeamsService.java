
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Service class for team management operations
 */
public class RandomTeamsService implements ITeamsService {
    private static final Logger logger = LoggerFactory.getLogger(RandomTeamsService.class);
    private static final int DEFAULT_TEAM_SIZE = 10;

    /**
     * Split students into teams of specified size
     * 
     * @param students List of all students
     * @param teamSize Size of each team (default: 10)
     * @return List of teams, each containing a list of students
     */
    public List<List<Student>> splitIntoTeams(List<Student> students, int teamSize) {
        if (students == null || students.isEmpty()) {
            logger.warn("No students provided for team creation");
            return new ArrayList<>();
        }

        if (teamSize <= 0) {
            teamSize = DEFAULT_TEAM_SIZE;
        }

        List<List<Student>> teams = new ArrayList<>();
        List<Student> shuffledStudents = new ArrayList<>(students);

        // Optional: Shuffle students for random team assignment
        Collections.shuffle(shuffledStudents);

        for (int i = 0; i < shuffledStudents.size(); i += teamSize) {
            int endIndex = Math.min(i + teamSize, shuffledStudents.size());
            List<Student> team = new ArrayList<>(shuffledStudents.subList(i, endIndex));
            int teamNo = i / teamSize;
            String teamName = "Team_" + teamNo;

            // Use stream to update team property
            team.stream()
                    .forEach(student -> student.setTeam(teamName));
            teams.add(team);
        }

        logger.info("Created {} teams from {} students (team size: {})",
                teams.size(), students.size(), teamSize);

        // Log team distribution
        for (int i = 0; i < teams.size(); i++) {
            logger.info("Team {} has {} students", i + 1, teams.get(i).size());
        }

        return teams;
    }

    /**
     * Split students into teams of default size (10)
     * 
     * @param students List of all students
     * @return List of teams, each containing a list of students
     */
    public List<List<Student>> splitIntoTeams(List<Student> students) {
        return splitIntoTeams(students, DEFAULT_TEAM_SIZE);
    }

    /**
     * Shuffle students randomly for team assignment
     * 
     * @param students List of students to shuffle
     * @return New list with shuffled students
     */
    public List<Student> shuffleStudents(List<Student> students) {
        List<Student> shuffled = new ArrayList<>(students);
        Collections.shuffle(shuffled);
        logger.info("Shuffled {} students for random team assignment", students.size());
        return shuffled;
    }

    /**
     * Get statistics about team distribution
     * 
     * @param teams List of teams
     * @return Team statistics as a formatted string
     */
    public String getTeamStatistics(List<List<Student>> teams) {
        if (teams.isEmpty()) {
            return "No teams created";
        }

        int totalStudents = teams.stream().mapToInt(List::size).sum();
        int minTeamSize = teams.stream().mapToInt(List::size).min().orElse(0);
        int maxTeamSize = teams.stream().mapToInt(List::size).max().orElse(0);
        double avgTeamSize = teams.stream().mapToInt(List::size).average().orElse(0.0);

        return String.format(
                "Team Statistics:\n" +
                        "- Total Teams: %d\n" +
                        "- Total Students: %d\n" +
                        "- Min Team Size: %d\n" +
                        "- Max Team Size: %d\n" +
                        "- Average Team Size: %.2f",
                teams.size(), totalStudents, minTeamSize, maxTeamSize, avgTeamSize);
    }
}