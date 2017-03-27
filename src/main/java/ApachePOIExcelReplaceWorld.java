import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Scanner;

public class ApachePOIExcelReplaceWorld {

    private static final String FILE_NAME = "MyFirstExcel.xlsx";
    private String valueCell;

    public void replaceWordExcel() {

        System.out.print("Enter word: ");
        Scanner scannerWord = new Scanner(System.in);
        String inputWord = scannerWord.nextLine();

        System.out.print("Enter new word: ");
        Scanner scannerNewWord = new Scanner(System.in);
        String inputNewWord = scannerNewWord.nextLine();

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
                        if (valueCell.equalsIgnoreCase(inputWord)) {
                            currentCell.setCellValue(inputNewWord);
                            found = true;
                        }
                    }   else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        valueCell = String.valueOf(currentCell.getNumericCellValue());
/*                        if (valueCell.equalsIgnoreCase(inputWord){
                            currentCell.setCellValue(inputNewWord);
*/                         found = true;
                    }
                }
            }
            if (found) {
                System.out.println("Word has been replaced");
            } else if (!found)
                System.out.println("Word has not been replaced");
            try {
                FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
                workbook.write(outputStream);
                workbook.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.out.println("Done");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
