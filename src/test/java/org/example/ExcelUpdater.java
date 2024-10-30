package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Random;


public class ExcelUpdater {

    public static void main(String[] args) {
        String filePath = System.getProperty("user.dir")+"/src/test/resources/Random_Excel.xlsx";
        String filePathAfterUpdate = System.getProperty("user.dir")+"/output/Random_Excel_Populated.xlsx";
        List<String> sportsList = Arrays.asList("Cricket","Football","Baseball","Shooting","Archery","Swimming","Badminton","Tennis","Hockey","Skating");


        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Iterator<Sheet> sheetIterator = workbook.sheetIterator();

            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();

                for (Row row : sheet) {
                    Cell firstCell = row.getCell(0);
                    if (firstCell == null || firstCell.getCellType() == CellType.BLANK) {
                        continue;
                    }

                    for (Cell cell : row) {
                        if (cell.getCellType() == CellType.FORMULA) {
                            continue;
                        }

                        CellStyle cellStyle = cell.getCellStyle();
                        if (isGreyCell(cellStyle) || isPeachCell(cellStyle)) {
                            continue;
                        }

                        if (cell.toString().contains("Please update this cell with the name of the revenue stream.")) {
                            cell.setCellValue(sportsList.stream().findFirst().get());
                           sportsList = sportsList.stream().skip(1).toList();
                            continue;
                        }

                        if (cell.getCellType() == CellType.BLANK) {
                            cell.setCellValue(getRandomInteger());
                        }

                    }
                }
            }
                FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            formulaEvaluator.evaluateAll();
                try (FileOutputStream fos = new FileOutputStream(filePathAfterUpdate)) {
                    workbook.write(fos);
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }


                System.out.println("Excel file updated successfully!");

            } catch(IOException e){
                e.printStackTrace();
            }

    }

    private static byte[] getRGBValue(CellStyle cellStyle){
        Color color = cellStyle.getFillForegroundColorColor();

        if (color instanceof XSSFColor) {
            XSSFColor xssfColor = (XSSFColor) color;
            return xssfColor.getRGB();
        }
        return  new byte[0];

    }

    private static boolean isGreyCell(CellStyle cellStyle) {
        if (cellStyle == null)
            return false;
        byte[] rgb = getRGBValue(cellStyle);
        return rgb != null && rgb.length > 0 && rgb[0] == (byte) -39 && rgb[1] == (byte) -39 && rgb[2] == (byte) -39;
    }

    private static boolean isPeachCell(CellStyle cellStyle) {
        if (cellStyle == null)
            return false;
        byte[] rgb = getRGBValue(cellStyle);
            return rgb != null && rgb.length > 0 && rgb[0] == (byte) -1 && rgb[1] == (byte) -24 && rgb[2] == (byte) -47;
    }

    private static Integer getRandomInteger(){
       return new Random().nextInt(1000,9999);
    }
}
