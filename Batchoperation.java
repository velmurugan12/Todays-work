package com.read_cell;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Batchoperation {

	public static void main(String[] args) throws IOException, SQLException
    {
        InputStream ExcelFileToRead = new FileInputStream("C:/Users/v.sudalaimani/Desktop/read.xlsx");
        XSSFWorkbook  wb = new XSSFWorkbook("C:/Users/v.sudalaimani/Desktop/read.xlsx");
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row; 
        XSSFCell cell;
        Iterator rows = sheet.rowIterator();
        
        String[] a = null;
    	int numberOfElements = 3;
    	a = new String[numberOfElements];

        while (rows.hasNext())
        {
            row=(XSSFRow) rows.next();
            Iterator cells = row.cellIterator();
            
           
            while (cells.hasNext())
            {
            	for(int i=0;i<3;i++){  
                cell=(XSSFCell) cells.next();   
                if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING)
                {
                	
                			a[i]=cell.getStringCellValue();
                		}
					
            }
        }      
    }
          
                		 try {
                				Class.forName("com.mysql.jdbc.Driver");

                				Connection con = DriverManager.getConnection("jdbc:mysql://127.0.0.1:3306/mydatabase", "root",
                						"vaticancameious");

                				Statement statement = con.createStatement();
                
                				statement.addBatch("INSERT INTO data(empname)  VALUES('"+a[0]+"');");
                				statement.addBatch("INSERT INTO data(empname)  VALUES('"+a[1]+"');");
                				statement.addBatch("INSERT INTO data(empname)  VALUES('"+a[2]+"');");
                                statement.executeBatch();
                	        }catch (Exception e) {
                				System.out.println(e);
                			}
                		 finally
                			{
                				
                				Connection con = null;
                				con.close();
                			}	
                
      }
 }

                      
        
    

