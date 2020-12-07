package com.github.nlobanov.sdc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;
import java.util.regex.Pattern;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.*;;

public class Main {

    public static void main(String[] args) throws FileNotFoundException {
        String filePath = getFilePath();
        FileInputStream inputStream = new FileInputStream(filePath);
        printNumericCells(inputStream);
    }

    public static String getFilePath() {
        String filePath = null;

        try (Scanner in = new Scanner(System.in)) {
            System.out.println("Enter exact file path:");
            filePath = in.nextLine();
            while(!new File(filePath).isFile()) {
                System.out.println("The document name or path is not valid. Try again:");
                filePath = in.nextLine();
            }
        } catch(Exception ex) {
            ex.printStackTrace();
        }

        return filePath;
    }

    public static void printNumericCells(FileInputStream inputStream) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream)) {
            XSSFSheet firstSheet = workbook.getSheetAt(0);

            for (int i = 0; i <= firstSheet.getRow(1).getLastCellNum(); i++) {
                for (int j = 0; j < firstSheet.getLastRowNum(); j++) {
                    XSSFRow row = firstSheet.getRow(j);
                    XSSFCell cell = row.getCell(i);
                    if(cell != null) {
                        printNumber(cell);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void printNumber(XSSFCell cell) {
        switch(cell.getCellType()) {
            case STRING:
                if(Pattern.matches("^[0-9]+$", cell.getStringCellValue()))
                    System.out.println(Integer.parseInt(cell.getStringCellValue()));
                break;
            case NUMERIC:
                System.out.println((int) cell.getNumericCellValue());
                break;
            default: break;
        }
    }
}
