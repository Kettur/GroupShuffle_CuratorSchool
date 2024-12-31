import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.commons.math3.util.Pair;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeOverrideDocument;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.JTextComponent;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;

import static java.lang.Math.floor;
import static java.lang.Math.round;


public class Main {
    private ArrayList<String> blockArray = new ArrayList<>();
    private ArrayList<Integer> blockGroupArray = new ArrayList<>();
    private ArrayList<String[]> blockTeacherArray = new ArrayList<>();
    private ArrayList<String[]> potoksArray;
    private ArrayList<ArrayList<String[]>> itogBlyatMatrixList = new ArrayList<>();
    private int numOfBlocks, numOfPotoks, count = 0, countBlocks = 0;
    private File chosenFile;
    private JFrame frameFirst;
    private JPanel panelFirst, panelSecond;
    private JTextField flowCountText, filePathText, blockCountText, peopleInBlockText, blockNameText, teachersInBlockText, blocksLeftCountText;
    private JTextField flowCount, filePath, blockCount, peopleInBlock, teachersInBlock, blockName;
    private JButton selectFile, nextButton, nextBlockButton;
    Main(){
        frameFirst = new JFrame("GroupDivider");
        frameFirst.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frameFirst.setLayout(new FlowLayout());

        panelFirst = new JPanel(new GridLayout(4,2));
        panelSecond = new JPanel(new GridLayout(4, 2));
        flowCountText = new JTextField("Количество потоков: ");
        flowCountText.setEditable(false);
        blockCountText = new JTextField("Количество блоков в потоке: ");
        blockCountText.setEditable(false);
        filePathText = new JTextField("Путь к исходному файлу: ");
        filePathText.setEditable(false);
        peopleInBlockText = new JTextField("Количество групп: ");
        peopleInBlockText.setEditable(false);
        teachersInBlockText = new JTextField("Список - аудитория;препод: ");
        teachersInBlockText.setEditable(false);
        blockNameText = new JTextField("Наименование блока: ");
        blockNameText.setEditable(false);
        blocksLeftCountText = new JTextField();
        blocksLeftCountText.setEditable(false);
        selectFile = new JButton("Выбрать исходный файл");
        nextButton = new JButton("Далее");
        nextBlockButton = new JButton("Далее");
        filePath = new JTextField();
        flowCount = new JTextField();
        blockCount = new JTextField();
        peopleInBlock = new JTextField();
        teachersInBlock = new JTextField();
        blockName = new JTextField();


        selectFile.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                selectFile();
            }
        });
        nextButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    startDivision();
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });
        nextBlockButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    divideFlows();
                } catch (IOException ex) {
                    throw new RuntimeException(ex);
                }
            }
        });

        panelFirst.add(filePathText);
        panelFirst.add(filePath);
        panelFirst.add(flowCountText);
        panelFirst.add(flowCount);
        panelFirst.add(blockCountText);
        panelFirst.add(blockCount);
        panelFirst.add(selectFile);
        panelFirst.add(nextButton);

        panelSecond.add(blockNameText);
        panelSecond.add(blockName);
        panelSecond.add(peopleInBlockText);
        panelSecond.add(peopleInBlock);
        panelSecond.add(teachersInBlockText);
        panelSecond.add(teachersInBlock);
        peopleInBlockText.setVisible(false);
        peopleInBlock.setVisible(false);
        teachersInBlockText.setVisible(false);
        teachersInBlock.setVisible(false);
        panelSecond.add(nextBlockButton);
        panelSecond.add(blocksLeftCountText);

        frameFirst.add(panelFirst);
        frameFirst.add(panelSecond);
        panelSecond.setVisible(false);
        frameFirst.pack();
        frameFirst.setVisible(true);
    }

    public void selectFile(){
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Tables", "xlsx");
        fileChooser.setFileFilter(filter);
        int retval = fileChooser.showSaveDialog(selectFile);
        if (retval == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            if (file != null) {
                filePath.setText(file.getAbsolutePath());
                chosenFile = file;
            }
        }
    }
    public void startDivision() throws IOException {
        FileInputStream inputStream = new FileInputStream(chosenFile);

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
        Sheet sheet = xssfWorkbook.getSheet("Лист1");
        ArrayList<String> peopleList = new ArrayList<>();
        for (Row row : sheet) {
            peopleList.add(row.getCell(0).toString());
        }
        Collections.shuffle(peopleList);

        numOfPotoks = Integer.parseInt(flowCount.getText());

        potoksArray = new ArrayList<>();
        int potokSize = round((float) peopleList.size() / numOfPotoks);
        for (int i = 0; i < numOfPotoks; i++){
            if (i + 1 == numOfPotoks){
                potoksArray.add(peopleList.subList(i * potokSize, peopleList.size()).toArray(new String[peopleList.size() - potokSize * i]));
            }
            else{
                potoksArray.add(peopleList.subList(i * potokSize, (i + 1) * potokSize).toArray(new String[potokSize]));
            }
        }

        numOfBlocks = Integer.parseInt(blockCount.getText());
        blocksLeftCountText.setText("Осталось  названий  блоков: " + numOfBlocks);
        panelSecond.setVisible(true);
        panelFirst.setVisible(false);
    }
    public void divideFlows() throws IOException {
        count++;
        if(count < numOfBlocks){
            blocksLeftCountText.setText("Осталось  названий  блоков: " + (numOfBlocks - count));
            String str = blockName.getText();
            blockArray.add(str);

            blockName.setText("");
        }
        else if(count == numOfBlocks){
            String str = blockName.getText();
            blockArray.add(str);

            blockName.setText("");
            peopleInBlockText.setVisible(true);
            peopleInBlock.setVisible(true);
            teachersInBlockText.setVisible(true);
            teachersInBlock.setVisible(true);
            blockName.setEditable(false);
            blockName.setText(blockArray.get(0) + "_Поток_1");
            blocksLeftCountText.setText("Осталось заполнить блоков: " + (numOfBlocks * numOfPotoks - countBlocks));
        }
        else if(countBlocks < ((numOfPotoks * numOfBlocks) - 1) && count > numOfBlocks){
            countBlocks++;

            blockName.setText(blockArray.get(countBlocks % numOfBlocks) + "_Поток_" + ((countBlocks/numOfBlocks) + 1));
            blocksLeftCountText.setText("Осталось заполнить блоков: " + (numOfBlocks * numOfPotoks - countBlocks));

            int peopleInBlockInt = Integer.parseInt(peopleInBlock.getText());
            blockGroupArray.add(peopleInBlockInt);
            String teachers = teachersInBlock.getText();
            blockTeacherArray.add(teachers.split(","));
            peopleInBlock.setText("");
            teachersInBlock.setText("");

        }
        else{

            int peopleInBlockInt = Integer.parseInt(peopleInBlock.getText());
            blockGroupArray.add(peopleInBlockInt);
            String teachers = teachersInBlock.getText();
            blockTeacherArray.add(teachers.split(","));
            peopleInBlock.setText("");
            teachersInBlock.setText("");

            String userHome = System.getProperty("user.home");
            userHome+="/Desktop";
            File theDir = new File(userHome, "Потоки");
            if (!theDir.exists()){
                theDir.mkdirs();
            }
            userHome+="/Потоки";
            for(int p = 0; p < numOfPotoks; p++){
                File theDir1 = new File(userHome, "Поток " + (p + 1));
                if (!theDir1.exists()){
                    theDir1.mkdirs();
                }
            }

            for (int j = 0; j < numOfPotoks; j++) {
                for (int g = 0; g < numOfBlocks; g++) {
                    shuffleArray(potoksArray.get(j));

                    File file = new File(userHome + "/Поток " + (j + 1), blockArray.get(g) + "_Поток_" + (j + 1) + ".xlsx");
                    Workbook workbook = new XSSFWorkbook();
                    ArrayList<String[]> blockGroups = new ArrayList<>();
                    Sheet groupSheet = workbook.createSheet("Распределение");
                    int groupSize = (int) floor((float) potoksArray.get(j).length / blockGroupArray.get(g + j * numOfBlocks));
                    int leftPeople = potoksArray.get(j).length % blockGroupArray.get(g + j * numOfBlocks);
                    int peopleAdded = 0;
                    int tupoiSchet = 0;
                    for (int i = 0; i < blockGroupArray.get(g + j * numOfBlocks); i++) {
                        if (i + leftPeople >= blockGroupArray.get(g + j * numOfBlocks)) {
                            String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize + tupoiSchet, ((i + 1) * groupSize) + tupoiSchet + 1);
                            tupoiSchet++;
                            for (int k = 0; k < temp.length; k++) {
                                Row row = groupSheet.createRow(k + peopleAdded);
                                Cell cell = row.createCell(0);
                                cell.setCellValue(temp[k] + ";" + blockTeacherArray.get(g + j * numOfBlocks)[i].replace("@", " "));
                            }
                            Sheet groupSheetTeachers = workbook.createSheet(blockTeacherArray.get(g + j * numOfBlocks)[i].replace("@", " "));
                            for (int k = 0; k < temp.length; k++) {
                                Row row = groupSheetTeachers.createRow(k);
                                Cell cell = row.createCell(0);
                                cell.setCellValue(temp[k]);
                            }

                            peopleAdded+= temp.length;
                        }
                        else {
                            String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize, (i + 1) * groupSize);
                            for (int k = 0; k < temp.length; k++) {
                                Row row = groupSheet.createRow(k + peopleAdded);
                                Cell cell = row.createCell(0);
                                cell.setCellValue(temp[k] + ";" + blockTeacherArray.get(g + j * numOfBlocks)[i].replace("@", " "));
                            }
                            Sheet groupSheetTeachers = workbook.createSheet(blockTeacherArray.get(g + j * numOfBlocks)[i].replace("@", " "));
                            for (int k = 0; k < temp.length; k++) {
                                Row row = groupSheetTeachers.createRow(k);
                                Cell cell = row.createCell(0);
                                cell.setCellValue(temp[k]);
                            }

                            peopleAdded+= temp.length;
                        }
                    }

//                    Попрытка автоматизировать весь процесс, вплоть до создания пакетки, но пока неудачная

//                    ArrayList<String> pizdecPeopleList = new ArrayList<>();
//                    for (Row row2 : groupSheet) {
//                        pizdecPeopleList.add(row2.getCell(0).toString());
//                    }
//                    Collections.sort(pizdecPeopleList);
//
//
//                    int wroted = 0;
//                    for (int suka = 1; suka < pizdecPeopleList.size(); suka++ ){
//                        if (!compLines(pizdecPeopleList.get(suka-1), pizdecPeopleList.get(suka))){
//                            ArrayList<String[]> matrix = new ArrayList<>();
//                            for (int jopa = 0; jopa < suka-wroted; jopa++){
//                                matrix.add(pizdecPeopleList.get(jopa + wroted).split(";"));
//                            }
//
//                            Collections.sort(matrix, new Comparator<String[]>() {
//                                @Override
//                                public int compare(String[] o1, String[] o2) {
//                                    return o1[2].compareTo(o2[2]);
//                                }
//                            });
//
//                            int hit = 0;
//                            int addedToItog = 0;
//                            for(int speack = 1; speack < matrix.size(); speack++){
//                                if(!compLinesClass(matrix.get(speack)[2], matrix.get(speack-1)[2])){
//                                    hit++;
//                                }
//                                if (hit == 2){
//                                    ArrayList<String[]> temp = new ArrayList<>();
//                                    for(int t = 0; t < speack - addedToItog; t++){
//                                        temp.add(matrix.get(t + addedToItog));
//                                    }
//                                    itogBlyatMatrixList.add(temp);
//                                    addedToItog+=speack;
//                                    hit = 0;
//                                }
//                                if (speack+1 == matrix.size()){
//                                    ArrayList<String[]> temp = new ArrayList<>();
//                                    for(int t = 0; t < speack - addedToItog + 1; t++){
//                                        temp.add(matrix.get(t + addedToItog));
//                                    }
//                                    itogBlyatMatrixList.add(temp);
//                                }
//                            }
//
//                            wroted+=(suka-wroted);
//                        }
//                        if (suka+1 == pizdecPeopleList.size()){
//                            ArrayList<String[]> matrix = new ArrayList<>();
//                            for (int jopa = 0; jopa < suka-wroted+1; jopa++){
//                                matrix.add(pizdecPeopleList.get(jopa + wroted).split(";"));
//                            }
//
//                            Collections.sort(matrix, new Comparator<String[]>() {
//                                @Override
//                                public int compare(String[] o1, String[] o2) {
//                                    return o1[2].compareTo(o2[2]);
//                                }
//                            });
//                            int hit = 0;
//                            int addedToItog = 0;
//                            for(int speack = 1; speack < matrix.size(); speack++){
//                                if(!compLinesClass(matrix.get(speack)[2], matrix.get(speack-1)[2])){
//                                    hit++;
//                                }
//                                if (hit == 2){
//                                    ArrayList<String[]> temp = new ArrayList<>();
//                                    for(int t = 0; t < speack - addedToItog; t++){
//                                        temp.add(matrix.get(t + addedToItog));
//                                    }
//                                    itogBlyatMatrixList.add(temp);
//                                    addedToItog+=speack;
//                                    hit = 0;
//                                }
//                                if (speack+1 == matrix.size()){
//                                    ArrayList<String[]> temp = new ArrayList<>();
//                                    for(int t = 0; t < speack - addedToItog + 1; t++){
//                                        temp.add(matrix.get(t + addedToItog));
//                                    }
//                                    itogBlyatMatrixList.add(temp);
//                                }
//                            }
//                        }
//                    }
//                    Sheet groupDividedSheet = workbook.createSheet("РаспределениеПакетка");
//                    int totalRow = 0;
//                    for (int zaebalo = 0; zaebalo < itogBlyatMatrixList.size(); zaebalo++){
//
//                        for (int ryad = 0; ryad < itogBlyatMatrixList.get(zaebalo).size(); ryad++){
//                            for (int stolb = 0; stolb < itogBlyatMatrixList.get(zaebalo).get(ryad).length; stolb++){
//
//
//                            }
//                        }
//                    }

                    FileOutputStream fileOutputStream = new FileOutputStream(file);
                    workbook.write(fileOutputStream);
                    fileOutputStream.close();
                }
            }
            System.exit(0);
        }
    }
    static void shuffleArray(String[] ar){
        Random rnd = new Random();
        for (int i = ar.length - 1; i > 0; i--){
            int index = rnd.nextInt(i + 1);
            String a = ar[index];
            ar[index] = ar[i];
            ar[i] = a;
        }
    }
//    static boolean compLines(String prev, String pres){
//        int curr = 0;
//        while(prev.charAt(curr) != ';' && pres.charAt(curr) != ';'){
//            if(prev.charAt(curr) != pres.charAt(curr)){
//                return false;
//            }
//            curr++;
//        }
//        if (!(prev.charAt(curr) == ';' && pres.charAt(curr) == ';')){
//            return false;
//        }
//        return true;
//    }
//    static boolean compLinesClass(String prev, String pres) {
//        int curr = 0;
//        if (prev.length() == pres.length()) {
//            while (curr < prev.length()) {
//                if (prev.charAt(curr) != pres.charAt(curr)) {
//                    return false;
//                }
//                curr++;
//            }
//            return true;
//        }
//        return false;
//    }



    public static void main(String[] args) throws IOException {

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Main();
            }
        });
    }
}