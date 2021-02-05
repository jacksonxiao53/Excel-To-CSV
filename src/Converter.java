import java.util.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Converter {
	
	/**
	 * Get all the file names from the Input folder
	 * @param folder - Input folder
	 * @return - a list of all file names within the Input folder
	 */
	public static List<String> getFiles(File folder){
		List<String> fileList = new ArrayList<>();
		for(File fileEntry : folder.listFiles()) {
			fileList.add(fileEntry.getName());
		}
		return fileList;
	}
	
	/**
	 * 
	 * @param sheet - a single sheet from excel file
	 * @param sheetName - the name of the sheet
	 * @param fileName - the name of the excel file
	 */
	public static void convertExcelToCSV(org.apache.poi.ss.usermodel.Sheet sheet, String sheetName, String fileName) {
        StringBuilder data = new StringBuilder();
        try {
            Iterator<Row> rowIterator = sheet.iterator(); //Iterator for all rows in the excel file
            Row row; //initialize row
            Cell cell; //initialize cell
            while (rowIterator.hasNext()) { //iterate through each row
                row = rowIterator.next(); //get the next row

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator(); //Iterator for all cells in the excel file
                while (cellIterator.hasNext()) { //iterate through each cell within a row

                    cell = cellIterator.next(); //get the next cell
                    CellType type = cell.getCellType(); //get the type of the cell
                    switch (type) {
                        case BOOLEAN:
                            data.append(cell.getBooleanCellValue() + ",");
                            break;
                        case NUMERIC:
                            data.append(cell.getNumericCellValue() + ",");
                            break;
                        case STRING:
                            data.append(cell.getStringCellValue() + ",");
                            break;
                        case BLANK:
                            data.append("" + ",");
                            break;
                        case FORMULA:
                        	data.append(cell.getNumericCellValue() + ",");
                        	break;
                        default:
                            data.append(cell + ",");

                    }
                }
                data.append("\r\n");
            }
            Files.write(Paths.get("Output/" + fileName + "_" + sheetName + ".csv"),
                data.toString().getBytes("UTF-8")); //write to a new csv file
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	
	public static void main(String[] args) {
		File folder = new File("Input/"); //Input folder
        List<String> fileList = getFiles(folder); //get all file names within the Input folder
        
        for(String file : fileList) { //iterate through each file in the file list
        	String filePath = String.format("Input/%s", file); 
            try (InputStream input = new FileInputStream(filePath)) {
                Workbook wb = WorkbookFactory.create(input);

                for (int i = 0; i < wb.getNumberOfSheets(); i++) { //iterate through each sheet in excel file and convert it into csv
                    convertExcelToCSV(wb.getSheetAt(i), wb.getSheetAt(i).getSheetName(), file.substring(0,file.length()-5));
                    
                }
                System.out.println("All sheets has been converted for file: " + file);
            } catch (Exception ex) {
                System.out.println(ex.getMessage());
            } 
        }
	}

}

