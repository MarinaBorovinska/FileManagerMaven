import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

public class FileManager {

    public void createFile(){
        List<String> lines = Arrays.asList("first", "second", "third");
        Path file = Paths.get("new_file.txt");
        try {
            Files.write(file, lines, Charset.forName("UTF-8"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("File new_file.txt is created");
    }

    public void deleteFile(){

        Scanner scanner = new Scanner(System.in);
        System.out.println("Enter File name: ");
        String inputPath = scanner.nextLine();
        try{
            File file = new File(inputPath + ".txt");
            if(file.delete()){
                System.out.println(file.getName() + " is deleted");
            }
            else{
                System.out.println("Delete operation is failed.");
            }
        }catch(Exception e){
            e.printStackTrace();
        }

    }

    public void renameFile(){
        File oldFile = new File("new_file.txt");
        File newFile = new File("old_file.txt");

        if(oldFile.renameTo(newFile)){
            System.out.println("Rename succesful");
        }else{
            System.out.println("Rename failed");
        }
    }

    public void findWord(){

        System.out.print("Enter file name: ");
        Scanner scannerFile = new Scanner(System.in);
        String inputPath = scannerFile.nextLine();

        System.out.print("Enter word: ");
        Scanner scannerWord = new Scanner(System.in);
        String inputWord = scannerWord.nextLine();

        File file = new File(inputPath + ".txt");

        Scanner scanner = null;
        try {
            scanner = new Scanner(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        if (scanner != null) {
            while (scanner.hasNextLine()) {
                String nextWord = scanner.next();
                if (nextWord.equalsIgnoreCase(inputWord)) {
                    System.out.println("Found word is " + nextWord);
                }
                else System.out.println("Word is not found");
                break;
            }
        }
        //      scannerFile.close();
        //       scannerWord.close();
    }

    public void replaceWord(){

        System.out.print("Enter file name: ");
        Scanner scannerFile = new Scanner(System.in);
        String inputPath = scannerFile.nextLine();

        System.out.print("Enter word: ");
        Scanner scannerWord = new Scanner(System.in);
        String inputWord = scannerWord.nextLine();

        System.out.print("Enter new word: ");
        Scanner scannerNewWord = new Scanner(System.in);
        String inputNewWord = scannerNewWord.nextLine();

        try {
            File file = new File(inputPath + ".txt");
            BufferedReader reader = new BufferedReader(new FileReader(file));
            String line = "", oldtext = "";
            while ((line = reader.readLine()) != null) {
                oldtext += line + "\r\n";
            }
            reader.close();

            String newWord = oldtext.replaceAll(inputWord, "" + inputNewWord);

            FileWriter writer = new FileWriter(inputPath + ".txt");
            writer.write(newWord);
            System.out.println("Word replaced");
            writer.close();

            //          scannerFile.close();
            //          scannerWord.close();
            //          scannerNewWord.close();
        }
        catch (IOException ioe)
        {
            ioe.printStackTrace();
        }
    }

}