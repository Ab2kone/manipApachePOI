package excelOperations;

import com.sun.org.apache.xpath.internal.operations.Bool;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

// Workoob --> Sheet --> Rows --> Cells
public class WritingExcelDemo {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        Object emptydata[][] = {
                {"EmpID", "Name", "Job"},
                {101, "David", "Engineer"},
                {102, "Smith", "Manager"},
                {103, "Scott", "Analyste"}
        };

        int rows = emptydata.length;
        int cols = emptydata[0].length;

        System.out.println(rows);
        System.out.println(cols);


        for (int r=0; r < rows; r++){

            XSSFRow row = sheet.createRow(r);

            for (int c = 0; c < cols; c++){
                XSSFCell cell = row.createCell(c);
                Object value = emptydata[r][c];

                if (value instanceof String){
                    cell.setCellValue((String) value);
                }

                if (value instanceof Integer){
                    cell.setCellValue((Integer) value);
                }

                if (value instanceof Boolean){
                    cell.setCellValue((Boolean) value);
                }
            }
        }

        String filePath = "./datafiles/employee.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        outputStream.close();

        System.out.println("employee.xlsx file written successfully.......");

    }
}
