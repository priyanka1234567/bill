package backup_billing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;


//import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class apachepoiexample {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		
		        try
		        {
		            FileInputStream file = new FileInputStream(new File("E:\\bcbilling\\123456.xlsx"));
		 
		            //Create Workbook instance holding reference to .xlsx file
		            XSSFWorkbook workbook = new XSSFWorkbook(file);
		 
		            //Get first/desired sheet from the workbook
		            XSSFSheet sheet = workbook.getSheetAt(0);
		            
		            XSSFWorkbook wb = new XSSFWorkbook(); 
		            
		            //Create a blank sheet
		            XSSFSheet s1 = wb.createSheet("sheet1");
		            Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
		            Integer count=1;
		            
		            data.put(count++, new Object[] {"REQUEST_NO", "MASTER", "NBU_BACKUP_ID"});
		 
		            //Iterate through each rows one by one
		            Iterator<Row> rowIterator = sheet.iterator();
		            while (rowIterator.hasNext()) 
		            {
		                Row row = rowIterator.next();
		                //For each row, iterate through all the columns
		                Iterator<Cell> cellIterator = row.cellIterator();
		               // if(cell.getStringCellValue().equals("master_id"))
                    	//	cell.getColumnIndex()+1);}
		                String master="MASTER_ID";
		                Integer nbu_backup_id=0;
		                if(row.getRowNum()==0)
		                	continue;
		                 
		                while (cellIterator.hasNext()) 
		                {
		                    Cell cell = cellIterator.next();
		                    //Check the cell type and format accordingly
		                    //System.out.print(cell.getCellType()+" ");
		                   
		                    if(cell.getCellType().toString().equals("STRING") && (cell.getColumnIndex()+1)==5 && !cell.getStringCellValue().equals("master_id")) {		                    	
		                    	System.out.print(cell.getStringCellValue()+" ");
		                    	master=cell.getStringCellValue();
		                    	}
		                    else if(cell.getCellType().toString().equals("NUMERIC") && (cell.getColumnIndex()+1)==7)
		                    	{
		                    	System.out.print(cell.getNumericCellValue()+" " );	
		                    	nbu_backup_id=(int) cell.getNumericCellValue();
		                    	}
		                 
		                }
		             
		                data.put(count++, new Object[] {"12345678", master,nbu_backup_id});
		                System.out.println("");
		            }
		            
		            Set<Integer> keyset = data.keySet();
		            int rownum = 0;
		            for (Integer key : keyset)
		            {
		                Row row = s1.createRow(rownum++);
		                Object [] objArr = data.get(key);
		                int cellnum = 0;
		                for (Object obj : objArr)
		                {
		                   Cell cell = row.createCell(cellnum++);
		                   if(obj instanceof String)
		                        cell.setCellValue((String)obj);
		                    else if(obj instanceof Integer)
		                        cell.setCellValue((Integer)obj);
		                }
		            }
		            
		            FileOutputStream fileOut = new FileOutputStream("E:\\\\bcbilling\\\\234567.xlsx"); 
		            wb.write(fileOut); 
		            fileOut.close();
		            file.close();
		        } 
		        catch (Exception e) 
		        {
		            e.printStackTrace();
		        }
		    }
	

}
