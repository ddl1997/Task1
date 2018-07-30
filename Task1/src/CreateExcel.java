import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

public class CreateExcel {
	
	public static void create_excel(String sql)
	{
		Connection conn = getConn();
		try {
			PreparedStatement ps = (PreparedStatement)conn.prepareStatement(sql);
			ResultSet rs = (ResultSet) ps.executeQuery();
			String[] titles;
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	private static Connection getConn() {
	    String driver = "com.mysql.jdbc.Driver";
	    String url = "jdbc:mysql://localhost:3306/test";
	    String username = "root";
	    String password = "123456";
	    Connection conn = null;
	    try {
	        Class.forName(driver); //classLoader,���ض�Ӧ����
	        conn = (Connection) DriverManager.getConnection(url, username, password);
	    } catch (ClassNotFoundException e) {
	        e.printStackTrace();
	    } catch (SQLException e) {
	        e.printStackTrace();
	    }
	    return conn;
	}
	
	

}
