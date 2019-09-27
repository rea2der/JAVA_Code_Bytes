import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MAIL_CSV_TO_ORACLE {
	

	public static void insertDB (String FilePath) throws ClassNotFoundException, SQLException, IOException
	{
		  
		 FileReader reader1 = new FileReader("MAIL_IGSM_TO_ORACLE_CONFIG");
		 Properties properties = new Properties();
		 properties.load(reader1);
		
		  Connection conn = null;
		  Class.forName(properties.getProperty("db.class"));
		  conn = DriverManager.getConnection(properties.getProperty("db.connection"),properties.getProperty("db.username"),properties.getProperty("db.password"));
		  PreparedStatement preparedStmtSummary = null;
		  ResultSet resultSet = null;
		  StringBuffer sb = new StringBuffer();
		  String query=properties.getProperty("db.query1");
		  DataFormatter formatter = new DataFormatter();
		 
		  String tablename = properties.getProperty("db.tablename");
		  String query2="Truncate table " + tablename;
		  PreparedStatement p4=conn.prepareStatement(query2);
		  p4.executeUpdate();
		  
		  File myFile = new File(FilePath);
		  FileInputStream fis = new FileInputStream(myFile);
		  XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
		  XSSFSheet mySheet = myWorkBook.getSheetAt(0);
		  
		  String query1 = "SELECT count(*) As Valure FROM user_tab_columns WHERE table_name = '" + tablename +"'" ;
		  System.out.println(query1);
		  
		  PreparedStatement p1=conn.prepareStatement(query1);
		  resultSet = p1.executeQuery();
		  ResultSetMetaData rsmd = p1.getMetaData();
	      resultSet.next();
	      String columDb = resultSet.getString("Valure");
	      System.out.println("Columns in Database Table is "+columDb);
		  
		  int columV = mySheet.getRow(0).getPhysicalNumberOfCells();
		  System.out.println("Columns in Excel is "+columV);
		  
	  
		  int k =0;

		  if (columV == Integer.parseInt(columDb))
		  {
			 
			  PreparedStatement p=conn.prepareStatement(query);

			  for (int i = 1; i < mySheet.getPhysicalNumberOfRows(); i++) 
              {
			
					 					System.out.println("----------------------------");
					 				
					 					Cell projrefid1 = mySheet.getRow(i).getCell(0) ;
										String projrefid = formatter.formatCellValue(projrefid1);
										p.setString(1, projrefid);
		
										
					 					Cell orderid1 = mySheet.getRow(i).getCell(1) ;
										String orderid = formatter.formatCellValue(orderid1);
										p.setString(2, orderid);
	
										
					 					Cell end_date1 = mySheet.getRow(i).getCell(2) ;
										String end_date = formatter.formatCellValue(end_date1);
										//System.out.println(end_date);
										if(end_date.isEmpty())
										{
											p.setString(3, end_date);
										}
										else
										{
											String[] parts = end_date.split("/");
											String part1 = parts[0];
											String part2 = parts[1];
											String part3 = parts[2];
		
											String final_date = part2 + "/" + part1 + "/20"+ part3;
											//System.out.println(final_date);
											p.setString(3, final_date);
										}
										
										
										
										Cell Held_Value1 = mySheet.getRow(i).getCell(3) ;
										String Held_Value = formatter.formatCellValue(Held_Value1);
										p.setString(4, Held_Value);
		
										
					 					Cell EOT_Submitted1 = mySheet.getRow(i).getCell(4) ;
										String EOT_Submitted = formatter.formatCellValue(EOT_Submitted1);
										p.setString(5, EOT_Submitted);
										
									
										Cell EOT_Approved1 = mySheet.getRow(i).getCell(5) ;
										String EOT_Approved = formatter.formatCellValue(EOT_Approved1);
										p.setString(6, EOT_Approved);
		
										
					 					Cell EOT_Rejected1 = mySheet.getRow(i).getCell(6) ;
										String EOT_Rejected = formatter.formatCellValue(EOT_Rejected1);
										p.setString(7, EOT_Rejected);
										
										Cell EOT_Recalled1 = mySheet.getRow(i).getCell(7) ;
										String EOT_Recalled = formatter.formatCellValue(EOT_Recalled1);
										p.setString(8, EOT_Recalled);
		
										
					 					Cell Supplier1 = mySheet.getRow(i).getCell(8) ;
										String Supplier = formatter.formatCellValue(Supplier1);
										p.setString(9, Supplier);
										

										ResultSet rs=p.executeQuery();
										rs.close();
									
										k=k+1;
										System.out.println("Total Row Inserted is "+k);
				}
			  
			   preparedStmtSummary = conn.prepareStatement("commit");
			   resultSet = preparedStmtSummary.executeQuery();
			
		}else
		{
			System.out.println("Count of Excel Columns are not as Database columns. Please check your file");
		}
		  
		  

		  
	}

	
	
	public static void main(String[] args) {
		try 
		{
			 FileReader reader = new FileReader("MAIL_IGSM_TO_ORACLE_CONFIG");
			 Properties properties = new Properties();
			 properties.load(reader);
			 reader.close();
			
			 insertDB(properties.getProperty("db.pathtofile"));
			 
		} catch (ClassNotFoundException | SQLException | IOException e)
		{
			e.printStackTrace();
		}

	}

}
