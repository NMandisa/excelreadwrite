package za.co.fynbos.excel;

import za.co.fynbos.excel.read.EmployeeExcelWrite;
/**
 *
 * @author NMandisa Mkhungo
 */
public class ExcelReadWriteApplication {

    private static final String FILE_NAME = System.getProperty("user.dir") + "/sample_data.xlsx";

    public static void main(String[] args) {
        EmployeeExcelWrite employeeExcelWrite = new EmployeeExcelWrite(FILE_NAME);
        
    }

}
