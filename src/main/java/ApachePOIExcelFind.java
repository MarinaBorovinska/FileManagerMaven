import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

public class ApachePOIExcelFind {
    private static final String FILE_NAME = "MyFirstExcel.xlsx";
    private String valueCell;

    public void findExcel() {

        System.out.print("Enter word: ");
        Scanner scannerWord = new Scanner(System.in);
        String inputWord = scannerWord.nextLine();

        try {
            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            boolean found = false;
            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();


                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();

                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        valueCell = currentCell.getStringCellValue();

                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        valueCell = String.valueOf(currentCell.getNumericCellValue());
                    }

                    if (valueCell.equalsIgnoreCase(inputWord)) {
                        System.out.println("Found word is " + valueCell);
                        found = true;
                    }

                }
            }
            if (!found)
            System.out.println("Word is not found");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
