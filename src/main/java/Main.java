import java.util.Scanner;

public class Main {

    public static void main(String[] args) {

        FileManager fileManager = new FileManager();
        ApachePOIExcelWrite apachePOIExcelWrite = new ApachePOIExcelWrite();
        ApachePOIExcelRead apachePOIExcelRead = new ApachePOIExcelRead();
        ApachePOIExcelFind apachePOIExcelFind = new ApachePOIExcelFind();
        Scanner scanner = new Scanner(System.in);

        System.out.println("1. Create File");
        System.out.println("2. Delete File");
        System.out.println("3. Rename File");
        System.out.println("4. Find Word");
        System.out.println("5. Replace Word");
        System.out.println("0. Exit");

        System.out.println("11. Create and write Excel file");
        System.out.println("12. Read Excel file");
        System.out.println("13. Find Word in Excel file");

        int f = 1;
        while (f == 1) {
            System.out.println("Input a number:");
            int inputNumber = 0;
            if(scanner.hasNextInt())
                inputNumber = scanner.nextInt();
            switch (inputNumber) {
                case 0:
                    f = 0;
                    break;
                case 1:
                    fileManager.createFile();
                    break;
                case 2:
                    fileManager.deleteFile();
                    break;
                case 3:
                    fileManager.renameFile();
                    break;
                case 4:
                    fileManager.findWord();
                    break;
                case 5:
                    fileManager.replaceWord();
                    break;
                case 11:
                    apachePOIExcelWrite.writeExcel();
                    break;
                case 12:
                    apachePOIExcelRead.readExcel();
                    break;
                case 13:
                    apachePOIExcelFind.findExcel();
                    break;
            }
        }

    }
}
