package excelOperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExcelFile2 {

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
         * Utilisation des itérators pour le lecture
         * de contenu
         */

        Iterator iterator = sheet.iterator();

        while (iterator.hasNext()){
            XSSFRow row = (XSSFRow) iterator.next();

            Iterator cellIterator = row.cellIterator();

            while (cellIterator.hasNext()){
                XSSFCell cell = (XSSFCell) cellIterator.next();

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
