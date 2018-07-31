import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class CreateExcel {
	
	public String create_excel(String sql)
	{
		WritableWorkbook workbook = null;
		String fileName = "temp.xls";
		File file = new File(fileName);
		try {
//			if (file.exists())
//			{
//				workbook = Workbook.createWorkbook(file, Workbook.getWorkbook(file));
//			}
//			else
//			{
//				file.createNewFile();
//				workbook = Workbook.createWorkbook(file);
//			}
			file.createNewFile();
			workbook = Workbook.createWorkbook(file);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		Connection conn = getConn();
		try {
			PreparedStatement ps = (PreparedStatement)conn.prepareStatement(sql);
			ResultSet rs = (ResultSet) ps.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			
			WritableSheet s = null;
			int sheetCount = workbook.getNumberOfSheets();
			s = workbook.createSheet("Sheet" + sheetCount, sheetCount);
			for (int i = 0; i < rsmd.getColumnCount(); i++)
			{
				Label label = new Label(i, 0, rsmd.getColumnName(i + 1));
                s.addCell(label);
			}
			int rowCount = 1;
			while(rs.next()) {
				for (int i = 0; i < rsmd.getColumnCount(); i++)
				{
					s.addCell(new Label(i, rowCount, rs.getString(i + 1)));
				}
				rowCount++;
			}
			
			workbook.write();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
        }
		
		return file.getAbsolutePath();
		
	}
	
	private static Connection getConn() {
	    String driver = "com.mysql.cj.jdbc.Driver";
	    String url = "jdbc:mysql://localhost:3306/test?useSSL=true&serverTimezone=GMT%2B8";
	    String username = "root";
	    String password = "123456";
	    Connection conn = null;
	    try {
	        Class.forName(driver); 
	        conn = (Connection) DriverManager.getConnection(url, username, password);
	    } catch (ClassNotFoundException e) {
	        e.printStackTrace();
	    } catch (SQLException e) {
	        e.printStackTrace();
	    }
	    return conn;
	}
	
	public static void main(String[] args)
	{
		System.out.println(new CreateExcel().create_excel("SELECT * FROM table_1;"));
	}

}
