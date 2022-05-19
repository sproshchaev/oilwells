package ru.ps;

import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Класс App содержит методы работы с MS Excel
 *
 * @version 1.0
 * @autor Sergey Proshchaev
 */
public class App {

    private static final Logger logger = LoggerFactory.getLogger(App.class);

    public static void main(String[] args) throws IOException {

        String fileInStr = "Отчет о работе нефтяных скважин (итоги).xlsx";
        String fileOutStr = "Свод добычи нефти и воды.xlsx";

        logger.info("Открытие файла " + fileInStr + "...");
        XSSFWorkbook myExcelBookIn = new XSSFWorkbook(new FileInputStream(fileInStr));
        XSSFSheet myExcelSheet = myExcelBookIn.getSheet("Лист1");

        logger.info("Создание файла с отчетом " + fileOutStr + "...");
        XSSFWorkbook myExcelBookOut = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream(fileOutStr);
        XSSFSheet sheetOut = myExcelBookOut.createSheet("Лист1");

        logger.info("Генерация шапки в " + fileOutStr + "...");
        creatingHeaderInReport(sheetOut);

        logger.info("Обработка файла " + fileInStr + "...");

        String oilFieldName = "";
        String oilFormationName = "";
        int rowCountSheetOut = 6;

        for (int i = myExcelSheet.getFirstRowNum(); i < myExcelSheet.getLastRowNum(); i++) {

            for (int j = 0; j < getLastCellNum(myExcelSheet, i); j++) {

                if (readFromCell(myExcelSheet, i, j).contains("Месторождение:")) {
                    oilFieldName = readFromCell(myExcelSheet, i, j).substring(15);
                }


                if (readFromCell(myExcelSheet, i, j).contains("Итого по пласту")) {

                    oilFormationName = readFromCell(myExcelSheet, i, j).substring(16);

                    writeToCell(sheetOut, rowCountSheetOut, 0, oilFieldName, 4);
                    writeToCell(sheetOut, rowCountSheetOut, 1, "МЭР", 4);
                    writeToCell(sheetOut, rowCountSheetOut, 2, oilFormationName, 4);

                    writeToCell(sheetOut, rowCountSheetOut, 3, "", 4);

                    writeToCell(sheetOut, rowCountSheetOut, 4, readFromCell(myExcelSheet, i, 1).toString(), 1);
                    writeToCell(sheetOut, rowCountSheetOut, 5, readFromCell(myExcelSheet, i, 2).toString(), 1);
                    writeToCell(sheetOut, rowCountSheetOut, 6, readFromCell(myExcelSheet, i, 3).toString(), 1);

                    writeToCell(sheetOut, rowCountSheetOut, 7, readFromCell(myExcelSheet, i, 4).toString(), 1);
                    writeToCell(sheetOut, rowCountSheetOut, 8, readFromCell(myExcelSheet, i, 5).toString(), 1);
                    writeToCell(sheetOut, rowCountSheetOut, 9, readFromCell(myExcelSheet, i, 6).toString(), 1);

                    writeToCell(sheetOut, rowCountSheetOut, 10, readFromCell(myExcelSheet, i, 10).toString(), 1);

                    rowCountSheetOut++;
                }

            }
        }

        logger.info("Закрытие ресурсов файла " + fileOutStr + "...");
        myExcelBookOut.write(fileOut);
        fileOut.close();

        logger.info("Закрытие ресурсов файла " + fileInStr + "...");
        myExcelBookIn.close();

    }

    /**
     * Метод writeToCell записывает значение в ячейку листа MS Excel
     *
     * @param - XSSFSheet excelSheet имя листа
     * @param - int row строка ячейки
     * @param - int column столбец ячейки
     * @param - String value значение, записываемое в ячейку
     * @param - byte valueType типа значения в ячейке: 1 - NUMBER, 2 - DATE, 4 - TEXT
     */
    private static void writeToCell(XSSFSheet excelSheet, int row, int column, String value, int valueType) {

        XSSFRow rowObj;
        XSSFCell cellObj;

        if (excelSheet.getRow(row) == null) {
            rowObj = excelSheet.createRow((short) row);
        } else {
            rowObj = excelSheet.getRow((short) row);
        }

        if (rowObj.getCell(column) == null) {
            cellObj = rowObj.createCell(column);
        } else {
            cellObj = rowObj.getCell(column);
        }

        switch (valueType) {
            case 1:
                cellObj.setCellType(CellFormatType.NUMBER.ordinal());
                cellObj.setCellValue(Double.valueOf(value));
                break;
            case 2:
                cellObj.setCellType(CellFormatType.DATE.ordinal());
                cellObj.setCellValue(value);
                break;
            case 4:
                cellObj.setCellType(CellFormatType.TEXT.ordinal());
                cellObj.setCellValue(value);
                break;
            default:
                cellObj.setCellType(CellFormatType.TEXT.ordinal());
                cellObj.setCellValue(value);
        }
    }

    /**
     * Метод readFromCell читает значение из ячейки листа MS Excel
     *
     * @param - XSSFSheet excelSheet имя листа
     * @param - int row строка ячейки
     * @param - int column столбец ячейки
     * @return - String значение ячейки в текстовом формате
     */
    private static String readFromCell(XSSFSheet excelSheet, int row, int column) {
        String resultReadFromCell = "null";

        XSSFRow rowObj = excelSheet.getRow(row);
        XSSFCell cellObj;

        if (rowObj != null) {
            cellObj = rowObj.getCell(column);
            if (cellObj != null) {
                resultReadFromCell = cellObj.toString();
            } else {
                resultReadFromCell = "";
            }
        }

        return resultReadFromCell;
    }

    /**
     * Метод getLastCellNum возвращает значение последней заполненной ячейки в строке листа MS Excel
     *
     * @param - XSSFSheet excelSheet имя листа
     * @param - int row строка ячейки
     * @return - int значение последней заполненной ячейки в строке row
     */
    private static int getLastCellNum(XSSFSheet excelSheet, int row) {

        if (excelSheet.getRow(row) != null) {
            return excelSheet.getRow(row).getLastCellNum();
        } else {
            return 0;
        }
    }

    /**
     * Метод creatingHeaderInReport формрует заголовки отчета листа MS Excel
     *
     * @param - XSSFSheet excelSheet имя листа
     */
    private static void creatingHeaderInReport(XSSFSheet excelSheet) {

        writeToCell(excelSheet, 0, 0, "Отчет Свод добычи нефти и воды", 4);
        writeToCell(excelSheet, 1, 0, "Дата отчета " + new SimpleDateFormat("dd-MM-yyyy HH:mm:ss").format(new Date()), 4);
        writeToCell(excelSheet, 2, 0, "", 4);

        writeToCell(excelSheet, 3, 0, "Месторождение", 4);
        writeToCell(excelSheet, 3, 1, "Подр.", 4);
        writeToCell(excelSheet, 3, 2, "Объект", 4);
        writeToCell(excelSheet, 3, 3, "Число действ. скважин", 4);
        writeToCell(excelSheet, 3, 4, "Добыча нефти", 4);
        writeToCell(excelSheet, 3, 7, "Добыча воды", 4);
        writeToCell(excelSheet, 3, 10, "Суток работы с начала года", 4);

        writeToCell(excelSheet, 4, 0, "", 4);
        writeToCell(excelSheet, 4, 1, "", 4);
        writeToCell(excelSheet, 4, 2, "", 4);
        writeToCell(excelSheet, 4, 3, "", 4);
        writeToCell(excelSheet, 4, 4, "За месяц", 4);
        writeToCell(excelSheet, 4, 5, "С начала года", 4);
        writeToCell(excelSheet, 4, 6, "С начала разработки", 4);
        writeToCell(excelSheet, 4, 7, "За месяц", 4);
        writeToCell(excelSheet, 4, 8, "С начала года", 4);
        writeToCell(excelSheet, 4, 9, "С начала разработки", 4);
        writeToCell(excelSheet, 4, 10, "", 4);

        writeToCell(excelSheet, 5, 0, "", 4);
        writeToCell(excelSheet, 5, 1, "", 4);
        writeToCell(excelSheet, 5, 2, "", 4);
        writeToCell(excelSheet, 5, 3, "(1)", 4);
        writeToCell(excelSheet, 5, 4, "(2)", 4);
        writeToCell(excelSheet, 5, 5, "(3)", 4);
        writeToCell(excelSheet, 5, 6, "(4)", 4);
        writeToCell(excelSheet, 5, 7, "(5)", 4);
        writeToCell(excelSheet, 5, 8, "(6)", 4);
        writeToCell(excelSheet, 5, 9, "(7)", 4);
        writeToCell(excelSheet, 5, 10, "(46)", 4);
        
    }

}
