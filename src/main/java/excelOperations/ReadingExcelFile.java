package excelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadingExcelFile {

    public static void main(String[] args) throws IOException {

        String excelFilePath = "./datafiles/sampledatasafety.xlsx";
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(excelFilePath);

        /**
         * Ici je veux acceder à la 2 feuille
         * je passe son numero d'index en paramètre
         */
        XSSFSheet sheet = workbook.getSheetAt(1);      // XSSFSheet sheet = workbook.getSheet("sheet1");


        /**
         * reccupération de la dernière
         * ligne et colonne du fichier
         */
        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();

        for (int r = 0; r <= rows; r++){

            XSSFRow row = sheet.getRow(r);

            for (int c = 0; c < cols; c++){
                XSSFCell cell = row.getCell(c);
               switch ( cell.getCellType()) {
                   case STRING: System.out.print(cell.getStringCellValue()); break;
                   case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
                   case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
               }

                System.out.print(" | ");
            }

            System.out.println();
        }


    }
}
