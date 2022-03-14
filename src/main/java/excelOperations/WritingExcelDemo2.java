package excelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

// Workoob --> Sheet --> Rows --> Cells
public class WritingExcelDemo2 {

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        ArrayList<Object[]> empdata = new ArrayList<Object[]>();
                empdata.add(new Object[]{"Empid", "Name", "Job"});
                empdata.add(new Object[]{101, "David", "Engineer"});
                empdata.add(new Object[]{102, "Smith", "Manager"});
                empdata.add(new Object[]{103, "John", "Developer"});
                empdata.add(new Object[]{104, "Scott", "Analyst"});



        int rownum = 0;
        for (Object[] emp: empdata){
            XSSFRow row = sheet.createRow(rownum++);
            int cellnum = 0;

            for (Object value: emp){

                XSSFCell cell = row.createCell(cellnum++);
                if (value instanceof String)
                    cell.setCellValue((String) value);

                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);

                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }


        String filePath = "./datafiles/employee2.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        outputStream.close();

        System.out.println("employee2.xlsx file written successfully.......");

    }
}
