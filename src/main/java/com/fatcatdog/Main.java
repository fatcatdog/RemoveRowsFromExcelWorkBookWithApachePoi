package com.fatcatdog;

import java.io.*;
import org.apache.poi.ss.usermodel.*;

public class Main {

    private static void removeEmptyRows(Sheet sheet) {
        boolean stop = false;
        boolean nonBlankRowFound;
        Row lastRow;
        Cell cell;

        while (stop == false) {
            nonBlankRowFound = false;
            lastRow = sheet.getRow(sheet.getLastRowNum());

            for (int c = 0; c <= lastRow.getLastCellNum(); c++) {
                cell = lastRow.getCell(c);
                if (cell != null && lastRow.getCell(c).getCellType() !=  CellType.BLANK) {
                    nonBlankRowFound = true;
                }
            }
            if (nonBlankRowFound == true) {
                stop = true;
            } else {
                sheet.removeRow(lastRow);
            }
        }
    }

    public static void printNumberOfRowsInAllSheetsInWorkbook(Workbook workbook){
        int numberOfSheets = workbook.getNumberOfSheets();

        for(int i = 0; i < numberOfSheets; i++) {
            System.out.println("Sheet: " + i + " # of rows: " + workbook.getSheetAt(i).getPhysicalNumberOfRows());
        }
    }

    public static void printRows(Workbook workbook){
        int numberOfSheets = workbook.getNumberOfSheets();

        for(int i = 0; i < numberOfSheets; i++) {
            System.out.println("Sheet num: " + i);

            for(int j = 0; j < workbook.getSheetAt(i).getLastRowNum(); j++) {
                System.out.println("Row: " + j);

                for(int x = 0; x < workbook.getSheetAt(i).getRow(j).getLastCellNum(); x++) {
                    System.out.println(workbook.getSheetAt(i).getRow(j).getCell(x));
                }

            }
        }
    }

    public static void main(String[] args) throws IOException {
        String fileName  = "";

        Workbook workbook = WorkbookFactory.create(new File(fileName));

        System.out.println("Before processing:");
        System.out.println();
        printNumberOfRowsInAllSheetsInWorkbook(workbook);

//        printRows(workbook);

        for (Sheet sheet : workbook) {
//            removeEmptyRows(sheet);
        }

        try {
            FileOutputStream fileOut = new FileOutputStream(fileName);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("After processing:");
        System.out.println();
        printNumberOfRowsInAllSheetsInWorkbook(workbook);
    }

}
