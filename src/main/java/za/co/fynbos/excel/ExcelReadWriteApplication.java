package za.co.fynbos.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *
 * @author NMandisa Mkhungo
 */
public class ExcelReadWriteApplication {

    private static final String FILE_NAME = System.getProperty("user.dir") + "/sample_data.xlsx";

    public static void main(String[] args) {

        System.out.println("-----You are here :" + System.getProperty("user.dir"));

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employee");
        Object[][] employeedata = {
            {"Name", "Surname", "BoD", "Employee Number"},
            {"Limbani Ejiroghene","Idowu", "13/06/1951", "562261"},
            {"Idir Faraji", "Kayode", "30/11/1962", "800597"},
            {"Imamu Chinedu", "Adebayo", "07/07/1965", "504991"},
            {"Seti Siddhartha", "Temitope", "08/06/1966", "616129"},
            {"Antigonos Rúni", "Ayodele", "04/08/1966", "660287"},
            {"Kweku Julius", "Arendse", "22/12/1984", "943410"},
            {"Raganhar Theoderich", "Idowu", "21/06/1987", "335292"},
            {"Diodoros Cosmas", "Ayodele", "16/10/1993", "749890"},
            {"Emmerich Ekkebert", "Adebayo", "03/09/1994", "625126"},
            {"Berhanu Arnviðr", "Abiodun", "04/11/1996", "908645"},};
        int rowNum = 0;
        System.out.println("----------------Creating excel---------");
        

        for (Object[] employee : employeedata) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : employee) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println(" Document Created.");
    }

}
