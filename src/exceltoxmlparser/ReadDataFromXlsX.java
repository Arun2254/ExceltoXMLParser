package exceltoxmlparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import exceltoxmlparser.ReadValuesFromExcel;

public class ReadDataFromXlsX {

	public static List<String> arrName = new ArrayList<String>();
	public static String targetfilename="PO";

	public static void main(String[] args) throws IOException, InterruptedException {
		String excelFilePath = "support/PO_Values.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook AddCatalog = null; 
        ReadValuesFromExcel revfrmexcel= new ReadValuesFromExcel();
        
        //Workbook workbook = new XSSFWorkbook(inputStream);
        //Find the file extension by splitting file name in substring  and getting only extension name
	    String fileExtensionName = excelFilePath.substring(excelFilePath.indexOf("."));
	    System.out.println(fileExtensionName);
	    //Check condition if the file is a .xls file or .xlsx file
	    if(fileExtensionName.equals(".xls")){
	        //If it is .xls file then create object of HSSFWorkbook class
	        AddCatalog = new HSSFWorkbook(inputStream);
	    }
	    else if(fileExtensionName.equals(".xlsx")){
	        //If it is .xlsx file then create object of XSSFWorkbook class
	    	AddCatalog = new XSSFWorkbook(inputStream); 
	    }
        
//        Sheet firstSheet = AddCatalog.getSheetAt(0);
//        Iterator<Row> iterator = firstSheet.iterator();
         
//        while (iterator.hasNext()) {
//            Row nextRow = iterator.next();
//            Iterator<Cell> cellIterator = nextRow.cellIterator();
//             
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                 
//                switch (cell.getCellType()) {
//                    case Cell.CELL_TYPE_STRING:
//                        System.out.print(cell.getStringCellValue());
//                        break;
//                    case Cell.CELL_TYPE_BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue());
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.print(cell.getNumericCellValue());
//                        break;
//                }
//                System.out.print(" - ");
//            }
//            System.out.println();
//        }
         
	  //Read sheet inside the workbook by its name
	    Sheet AddCatalogSheet = AddCatalog.getSheetAt(0);
	    //Find number of rows in excel file
	    int rowcount = AddCatalogSheet.getLastRowNum()- AddCatalogSheet.getFirstRowNum();
	    System.out.println("Total row number: "+rowcount);
	    for(int i=1; i<rowcount+1; i++){
	        //Create a loop to get the cell values of a row for one iteration
	        Row row = AddCatalogSheet.getRow(i);
	        for(int j=0; j<row.getLastCellNum(); j++){
	            // Create an object reference of 'Cell' class
	            Cell cell = row.getCell(j);
	            // Add all the cell values of a particular row
	            arrName.add(cell.getStringCellValue());
	            }
	        revfrmexcel.generatePOXML(arrName);//targetfilename+i+".xml");
	        //System.out.println(arrName.get(0));
	        arrName.clear();
	        // Create an iterator to iterate through the arrayList- 'arrName'
	        Iterator<String> itr = arrName.iterator();
	        while(itr.hasNext()){
	            System.out.println(itr.next());
	        }
	        System.out.println(i);
       inputStream.close();
  
	}
	}
}
