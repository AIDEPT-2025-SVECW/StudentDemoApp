import java.util.List;

public interface ITeamsService {
	
	 public List<List<Student>> splitIntoTeams(List<Student> students, int teamSize);
	 
	 public List<List<Student>> splitIntoTeams(List<Student> students);
	 
	 public String getTeamStatistics(List<List<Student>> teams);
	 

}
