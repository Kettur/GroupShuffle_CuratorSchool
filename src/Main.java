import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.commons.math3.util.Pair;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.JTextComponent;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.*;

import static java.lang.Math.round;



public class Main {
    private ArrayList<Pair<String, Integer>> blockAndNumOfGroups = new ArrayList<>();
    private ArrayList<String[]> potoksArray;
    private int numOfBlocks, numOfPotoks, count = 0;
    private File chosenFile;
    private JFrame frameFirst;
    private JPanel panelFirst, panelSecond;
    private JTextField flowCountText, filePathText, blockCountText, peopleInBlockText, blockNameText, blocksLeftCountText;
    private JTextField flowCount, filePath, blockCount, peopleInBlock, blockName;
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
                potoksArray.add(peopleList.subList(i * potokSize, peopleList.size()).toArray(new String[peopleList.size() - potokSize * (numOfPotoks - 1)]));
            }
            else{
                potoksArray.add(peopleList.subList(i * potokSize, (i + 1) * potokSize).toArray(new String[potokSize]));
            }
        }

        numOfBlocks = Integer.parseInt(blockCount.getText());
        blocksLeftCountText.setText("Осталось блоков: " + numOfBlocks);
        panelSecond.setVisible(true);
        panelFirst.setVisible(false);
    }
    public void divideFlows() throws IOException {
        count++;
        blocksLeftCountText.setText("Осталось блоков: " + (numOfBlocks - count));
        if(count < numOfBlocks){
            String str = blockName.getText();
            int i1 = Integer.parseInt(peopleInBlock.getText());
            blockAndNumOfGroups.add(new Pair<>(str, i1));

            blockName.setText("");
            peopleInBlock.setText("");
        }
        else{
            String str = blockName.getText();
            int i1 = Integer.parseInt(peopleInBlock.getText());
            blockAndNumOfGroups.add(new Pair<>(str, i1));

            for (int j = 0; j < potoksArray.size(); j++) {
                for (Pair<String, Integer> block : blockAndNumOfGroups) {
                    shuffleArray(potoksArray.get(j));
                    String userHome = System.getProperty("user.home");
                    userHome+="/Desktop";
                    File file = new File(userHome, block.getFirst() + "_Поток_" + (j + 1) + ".xlsx");
                    Workbook workbook = new XSSFWorkbook();
                    ArrayList<String[]> blockGroups = new ArrayList<>();
                    int groupSize = round((float) potoksArray.get(j).length / block.getSecond());
                    for (int i = 0; i < block.getSecond(); i++) {
                        if (i + 1 == block.getSecond()) {
                            String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize, potoksArray.get(j).length);
                            Sheet groupSheet = workbook.createSheet("Группа " + (i + 1));
                            for (int k = 0; k < temp.length; k++) {
                                Row row = groupSheet.createRow(k);
                                Cell cell = row.createCell(0);
                                cell.setCellValue(temp[k]);
                            }
                        } else {
                            String[] temp = Arrays.copyOfRange(potoksArray.get(j), i * groupSize, (i + 1) * groupSize);
                            Sheet groupSheet = workbook.createSheet("Группа " + (i + 1));
                            for (int k = 0; k < temp.length; k++) {
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


    public static void main(String[] args) throws IOException {

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Main();
            }
        });
    }
}