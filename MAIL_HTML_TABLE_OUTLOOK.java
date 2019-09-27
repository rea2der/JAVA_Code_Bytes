import java.io.*;

import java.sql.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.script.ScriptException;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Locale;
import java.util.Properties;

//Main Code
	
    String host= "**";  
	final String user= "**";
	
	final String password="";  
	
	//Get the session object  
    Properties props = new Properties();  
    props.put("mail.smtp.host",host);  
    props.put("mail.smtp.auth", "true");
    
   Session session = Session.getDefaultInstance(props,  new javax.mail.Authenticator() 
   {  
	   protected PasswordAuthentication getPasswordAuthentication()
	   {  
		return new PasswordAuthentication(user,password);  
	   }  
   });  
   
  //Compose the message  
   try {  
    MimeMessage message = new MimeMessage(session); 
    YearMonth thisMonth    = YearMonth.now();
	DateTimeFormatter monthYearFormatter = DateTimeFormatter.ofPattern("MMM yy");
		  
    LocalDateTime ldt = LocalDateTime.now();
    message.setFrom(new InternetAddress(user));  
     message.setSubject("Status Report : "+ DateTimeFormatter.ofPattern("dd-MM-yyyy", Locale.ENGLISH).format(ldt)); 

    Connection conn = null;
    Class.forName("oracle.jdbc.OracleDriver");
    conn = DriverManager.getConnection("**");
    													
	PreparedStatement preparedStmtSummary = null;
	ResultSet resultSet = null;
  
	String manageraddress;
	ArrayList<String> metricname = Metric_Name;
    
	StringBuffer sb = new StringBuffer();
	String query2="";
	PreparedStatement p2=conn.prepareStatement(query2);
    p2.setString(1, Managername);
	
	ResultSet rs=p2.executeQuery();
    rs.next();
		
         manageraddress = rs.getString("");
         System.out.println("Manager Address is "+ manageraddress );
         String toaddress= "";
         String toaddress1= "";
         String toaddress2= "";
         String toaddress3= "";
         String toaddress4= "";
         message.addRecipient(Message.RecipientType.TO,new InternetAddress(manageraddress));
         message.addRecipient(Message.RecipientType.CC,new InternetAddress(toaddress1));
         message.addRecipient(Message.RecipientType.CC,new InternetAddress(toaddress2));
         message.addRecipient(Message.RecipientType.CC,new InternetAddress(toaddress3));
         message.addRecipient(Message.RecipientType.CC,new InternetAddress(toaddress4));
                      

		String query1 = "";		
	
   	
	 PreparedStatement p1=conn.prepareStatement(query1);
     p1.setString(1, metricname.get(0));
	 p1.setString(2, Managername);
		       
    resultSet = p1.executeQuery();

    ResultSetMetaData rsmd = p1.getMetaData();
	int columnCount = rsmd.getColumnCount();
	
	
    sb.append("Hi Team,");
	sb.append("<br>");
	sb.append("<br>");
	sb.append("<b>Please find the details below,</b>");
	sb.append("<br>");
	sb.append("<br>");
	sb.append("<table width=\"100%\" border=\"0\";style=\"font-family:calibri;\">");
	sb.append("<tr style=\"background-color: #337ab7; color: white;\">");
	SimpleDateFormat sdf = new SimpleDateFormat("MMMM-yy");	
	
	  for (int i = 1; i <= columnCount; i++)
	{
  		    		  
  		  if(rsmd.getColumnName(i).equalsIgnoreCase("METRIC_NAME")) {
		    	
          	 sb.append("<b><td width=\"30%\">"+"METRIC NAME"+"</td></b>");
		    }
  		  else if(rsmd.getColumnName(i).equalsIgnoreCase("")) {
	    	
         	 sb.append("<b><td width=\"15%\">"+" "+"</td></b>");
		    }
  		  else if(rsmd.getColumnName(i).equalsIgnoreCase("")) {
 	    	
         	 sb.append("<b><td width=\"15%\">"+" "+"</td></b>");
		    }
  		  else if(rsmd.getColumnName(i).equalsIgnoreCase("")) {
  	    	
         	 sb.append("<b><td width=\"20%\">"+""+"</td></b>");
		    }
  		 else if(rsmd.getColumnName(i).equalsIgnoreCase("")) {
  	    	
        	 sb.append("<b><td width=\"10%\">"+""+"</td></b>");
		    }
  		  else {
 	    	
  			sb.append("<b><td align=left>"+rsmd.getColumnName(i)+"</td></b>");
		    }
  		
	}
    
	 sb.append("</tr>");
		
	    int count=0;
       
       while (resultSet.next())
       {
          count ++;
          if(count %2==0)
          {
                 sb.append("<tr style=\"background-color: #dbe9f9; color: black\">");
                 for(int i = 1; i <= columnCount; i++)
                              {
                                     sb.append("<td align=\"left\" height=\"10px\">");
                                     sb.append(resultSet.getString(i));
                                     sb.append("</td>");
                              }
                 sb.append("</tr>");  
          }
          else
          {
                 sb.append("<tr>");             
                 for(int i = 1; i <= columnCount; i++)
                              {
                                     sb.append("<td align=\"left\" height=\"10px\">");
                                     sb.append(resultSet.getString(i));
                                     sb.append("</td>");
                              }
         
                 sb.append("</tr>");        
          }
       }
	sb.append("</table>");
}