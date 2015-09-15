package connections;

import java.beans.Statement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.util.ArrayList;

public class ExcelConnections {

	public Connection conn = null;
	//statements object used to run queries 
		public Statement stmnt = null; 
		//string object used to create queries
		public String testCaseQuery, DQuery;
		// ResultSet Object to hold the Query Result
		public ResultSet rs_Lable;
		// Metadata Object is created to
	
		
		public ArrayList<String>  get_ScriptIDs( String excelWorkbookName,String sheetName) throws Exception
		{
			ArrayList<String> testCasesNames = new ArrayList<String>();
			connectToExcel(excelWorkbookName+".xls");
			testCaseQuery = "Select TCID From ["+sheetName +"$] where Execute='Y'";
			rs_Lable = ((java.sql.Statement) stmnt).executeQuery(testCaseQuery);
			rsdata_Table= rs_Lable.getMetaData(); 
			System.out.println("rsdata size is: "+ rs_Lable.getRow());

			String colName = "";
			while(rs_Lable.next())
			{
				if(rs_Lable.getString("TCID")!=null)
				{
					testCasesNames.add(rs_Lable.getString("TCID"));
				}
			}
				
			return testCasesNames;
		}
	
	
	public void connectToExcel(String excelName) throws Exception
	{
		Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");

		//ITestProject myRftProj = getCurrentProject();  
		//String projLocation = myRftProj.getLocation();   
//		String exPath = projLocation +"\\Data\\"+ excelName ;

		//System.out.println("Path of the Excel :"+exPath);

		conn = DriverManager.getConnection("jdbc:odbc:Driver={Microsoft Excel Driver (*.xls)};DBQ=" + exPath + "; READONLY=false");
		System.out.println("connection : "+conn.getCatalog());

		stmnt = (Statement) conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
		System.out.println("statement : "+stmnt);

	}
	
	
}
