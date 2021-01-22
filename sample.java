import java.io.File;
import java.io.FileOutputStream;

import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class demo {

	public static void main(String[] args) {
		
		 XSSFWorkbook workbook = new XSSFWorkbook(); 
		  
	        // Create a blank sheet 
	        
	        Scanner sc=new Scanner(System.in);
	        System.out.println("enter the excel sheet name");
	        String s=sc.nextLine();
	        XSSFSheet sheet = workbook.createSheet(s); 
	        		
	  
	        // This data needs to be written (Object[]) 
	        Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>(); 
	        System.out.println("enter the column counts");
	        int n;
	        n=sc.nextInt();
	        System.out.println("enter the column names");
	        
	        String s2,s3,s4;
	        s2=sc.nextLine();
	        s3=sc.nextLine();
	        s4=sc.nextLine();
	        
	        data.put(1, new Object[]{ s2, s3, s4}); 
	        String s5,s6,s7;
	        int j=0;
	        for(int i=0;i<n;i++)
	        {
	        	s5=sc.nextLine();
	        	s6=sc.nextLine();
	        	s7=sc.nextLine();
	        	 data.put(j, new Object[]{ s5,s6,s7}); 
	        	 j++;
	        }
	       
	        Set<Integer> keyset = data.keySet(); 
	        int rownum = 0; 
	        for (Integer key : keyset) { 
	          
	            Row row = sheet.createRow(rownum++); 
	            Object[] objArr = data.get(key); 
	            int cellnum = 0; 
	            for (Object obj : objArr) {  
	                Cell cell = row.createCell(cellnum++); 
	                if (obj instanceof String) 
	                    cell.setCellValue((String)obj); 
	                else if (obj instanceof Integer) 
	                    cell.setCellValue((Integer)obj); 
	            } 
	        } 
	        try { 
	           
	            FileOutputStream out = new FileOutputStream(new File("f:\\demo.xlsx")); 
	            workbook.write(out); 
	            out.close(); 
	            System.out.println("demo.xlsx written successfully on disk."); 
	        } 
	        catch (Exception e) { 
	            e.printStackTrace(); 
	        } 
	  
	}

}
