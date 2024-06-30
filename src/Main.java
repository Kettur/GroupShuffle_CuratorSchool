import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import static java.lang.Math.round;



public class Main {

    static void shuffleArray(String[] ar){
        Random rnd = new Random();
        for (int i = ar.length - 1; i > 0; i--){
            int index = rnd.nextInt(i + 1);
            String a = ar[index];
            ar[index] = ar[i];
            ar[i] = a;
        }
    }


    public static void main(String[] args) throws IOException {
        Scanner sc = new Scanner(System.in);
        String path = "C:\\Users\\Кирилл\\Desktop\\dataTest.xlsx";
        FileInputStream inputStream = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
        Sheet sheet = xssfWorkbook.getSheet("Лист1");
        ArrayList<String> peopleList = new ArrayList<>();
        for (Row row : sheet) {
            peopleList.add(row.getCell(0).toString());
        }
        Collections.shuffle(peopleList);
        System.out.println("Введите кол-во потоков ");
        int numOfPotoks = sc.nextInt();
        ArrayList<String[]> potoksArray = new ArrayList<>();
        int potokSize = round((float) peopleList.size() / numOfPotoks);
        for (int i = 0; i < numOfPotoks; i++){
            if (i + 1 == numOfPotoks){
                potoksArray.add(peopleList.subList(i * potokSize, peopleList.size()).toArray(new String[peopleList.size() - potokSize * (numOfPotoks - 1)]));
            }
            else{
                potoksArray.add(peopleList.subList(i * potokSize, (i + 1) * potokSize).toArray(new String[potokSize]));
            }
        }

        System.out.println("Введите количство блоков");
        int numOfBlocks = sc.nextInt();
        ArrayList<Pair<String, Integer>> blockAndNumOfGroups = new ArrayList<>();
        for (int i = 0; i < numOfBlocks; i++){
            System.out.println("Введите название блока и количество групп");
            String str = sc.next();
            int i1 = sc.nextInt();
            blockAndNumOfGroups.add(new Pair<>(str, i1));
        }

        for (int j = 0; j < potoksArray.size(); j++){
            for (Pair<String, Integer> block: blockAndNumOfGroups){
                shuffleArray(potoksArray.get(j));
                File file = new File("C:\\Users\\Кирилл\\Desktop\\" + block.getFirst() + "_Поток_" + (j + 1) + ".xlsx");
                Workbook workbook = new XSSFWorkbook();

                ArrayList<String[]> blockGroups = new ArrayList<>();
                int groupSize = round((float) potoksArray.get(j).length / block.getSecond());
                for (int i = 0; i < block.getSecond(); i++){
                    if (i + 1 == block.getSecond()){
                        String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize, potoksArray.get(j).length);
                        Sheet groupSheet = workbook.createSheet("Группа " + (i + 1));
                        for (int k = 0; k < temp.length; k++){
                            Row row = groupSheet.createRow(k);
                            Cell cell = row.createCell(0);
                            cell.setCellValue(temp[k]);
                        }
                    }
                    else{
                        String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize, (i + 1) * groupSize);
                        Sheet groupSheet = workbook.createSheet("Группа " + (i + 1));
                        for (int k = 0; k < temp.length; k++){
                            Row row = groupSheet.createRow(k);
                            Cell cell = row.createCell(0);
                            cell.setCellValue(temp[k]);
                        }
                    }
                }
                FileOutputStream fileOutputStream = new FileOutputStream(file);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
            }
        }
    }
}