package za.co.fynbos.excel.read;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author NMandisa Mkhungo
 */
public class EmployeeExcelWrite {

    public EmployeeExcelWrite(String filename) {
        System.out.println("**********************************************");
        System.out.println("You are here :" + System.getProperty("user.dir"));
        System.out.println("**********************************************");

        Object[][] employeedata = {
            {"Name", "Surname", "BoD", "Employee Number"},
            {"Limbani Ejiroghene", "Idowu", "13/06/1951", "562261"},
            {"Idir Faraji", "Kayode", "30/11/1962", "800597"},
            {"Imamu Chinedu", "Adebayo", "07/07/1965", "504991"},
            {"Seti Siddhartha", "Temitope", "08/06/1966", "616129"},
            {"Antigonos Rúni", "Ayodele", "04/08/1966", "660287"},
            {"Kweku Julius", "Arendse", "22/12/1984", "943410"},
            {"Raganhar Theoderich", "Idowu", "21/06/1987", "335292"},
            {"Diodoros Cosmas", "Ayodele", "16/10/1993", "749890"},
            {"Emmerich Ekkebert", "Adebayo", "03/09/1994", "625126"},
            {"Berhanu Arnviðr", "Abiodun", "04/11/1996", "908645"},};

        XSSFWorkbook workbook = new XSSFWorkbook();//Creating an instance of a workbook
        XSSFSheet sheet = workbook.createSheet("Employee");//Creating an instance of a sheet 

        int rowNum = 0;//Initializing the Row
        for (int i = 0; i < employeedata.length; i++) {//Looping through the 2D Array of Objects
            Row row = sheet.createRow(rowNum++);//Creating the Row in the desired sheet.
            int ColNum = 0;//Initializa the Column Right after the Row has been created.
            for (int j = 0; j < employeedata[i].length; j++) {
                Cell cell = row.createCell(ColNum++);//create a Cell for each Column Row in desired sheet
                if (employeedata[i][j] != null) { //Checking weather the field is an String/Integer 
                    cell.setCellValue(employeedata[i][j].toString());
                    System.out.print("|" + employeedata[i][j].toString() + " | " + " \t" + " \t");
                }
            }
            System.out.println("");
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(filename);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(" Document Created Successfully.");

    }
}
