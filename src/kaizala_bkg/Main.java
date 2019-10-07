package kaizala_bkg;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;


public class Main {
	
	public static final String SAMPLE_XLSX_FILE_PATH = "./Tariff.xls";
	public static HashMap<String, Set<Long>> en_dict = new HashMap<String , Set<Long>>();
	
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		// Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        //System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
                
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });


        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();
        
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        int rowIt = 1;
        boolean header = true;
        while(/*rowIterator.hasNext()*/ rowIt>=0) {
        	rowIt--;
        	Row row = rowIterator.next();
        	
        	// Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();
            
            while(cellIterator.hasNext()) {
            	Cell cell = cellIterator.next();
            	String HRMNZD_CODE = dataFormatter.formatCellValue(cell);
            	String ITEM_ARBC_DESC = null , ITEM_ENG_DESC = null;
            	if(cellIterator.hasNext()) {
            		cell = cellIterator.next();
            		ITEM_ARBC_DESC = dataFormatter.formatCellValue(cell);
            		if(cellIterator.hasNext()) {
            			cell = cellIterator.next();
            			ITEM_ENG_DESC = dataFormatter.formatCellValue(cell);
            		}
                	
            	}
            	if(header == true) {
            		header = false;
            		continue;
            	}
            	Long code = Long.parseLong(HRMNZD_CODE);
            	String []partes_ar_desc = ITEM_ARBC_DESC.split(" ");
            	String []partes_en_desc = ITEM_ENG_DESC.split(" ");
            	
            	for(int i = 0; i<partes_en_desc.length ; i++) {
            		//String part = clearStrings(partes_en_desc[i]);
            		String part = partes_en_desc[i];
            		//System.out.println(part);
            		if(en_dict.containsKey(part)) {
            			en_dict.get(part).add(code);
            		}else {
            			en_dict.put(part , new HashSet<Long>());
            			en_dict.get(part).add(code);
            		}
            	}
            	
            	for (String key : en_dict.keySet())  
            		System.out.println(key + ":" + en_dict.get(key).toString());
            }
        }

        // Closing the workbook
        workbook.close();
        
        
	}
	
	public static String clearStrings(String str) {
		String temp = "";
		for(int i =0; i<str.length(); i++) {
			if( ( (int)str.charAt(i) >= 32 && (int)str.charAt(i)<= 64 )  || ((int)str.charAt(i) >=91 && (int)str.charAt(i) <= 96 ) || ((int)str.charAt(i) <= 123) || (int)str.charAt(i) == 1548 || (int)str.charAt(i) == 1563){
				continue;
			}
			temp += str.charAt(i);
		}
		return temp;
	}

}
