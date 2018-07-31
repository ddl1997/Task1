//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.InputStream;
import jxl.Sheet;
import jxl.Workbook;


public class ExcelToJson {
	
	public static String excel_to_json(InputStream input)
	{
		if (input == null)
		{
			return "";
		}
		String result  = "{ ";
		Workbook wb = null;
		try {
			wb = Workbook.getWorkbook(input);
			int sheet_size = wb.getNumberOfSheets();
			for (int i = 0; i < sheet_size; i++)
			{
				
				Sheet s = wb.getSheet(i);
				result += "\"" + s.getName() + "\" : [ ";
				for (int j = 0; j < s.getRows(); j++)
				{
					result += "{ ";
					for (int k = 0; k < s.getColumns(); k++)
					{
						String cellinfo = s.getCell(k, j).getContents();
						result += "\"" + k + "\" : \"" + cellinfo + "\" ";
						if (k < s.getColumns() - 1)
						{
							result += ", ";
						}
					}
					result = j == s.getRows() - 1 ? result + "} " : result + "}, ";
				}
				result = i == sheet_size - 1 ? result + "] " : result + "], ";
			}
			result += "}";
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null)
				wb.close();
		}
		
		return result;
	}
	
//	public static void main(String[] args)
//	{
//		File file = new File("D:"+ File.separator +"1.xls");
//
//		String excelpath = file.getAbsolutePath();
//		try {
//			InputStream is = new FileInputStream(excelpath);
//			System.out.println(excel_to_json(is));
//		} catch (FileNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//        
//	}

}
