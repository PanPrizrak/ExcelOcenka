/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.excelocenka;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
 
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
 
public class ExcelParser {
 
    public static String parse() throws FileNotFoundException {
    //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        JFileChooser window = new JFileChooser();
        int returnValue = window.showOpenDialog(null);
        inputStream = null;
        if(returnValue==JFileChooser.APPROVE_OPTION)
            inputStream = new FileInputStream(window.getSelectedFile());
            JOptionPane.showMessageDialog(null, window.getSelectedFile().toString());
        try {
            //inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
     //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(3);
        Iterator<Row> it = sheet.iterator();
     //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
      //перебираем возможные типы ячеек
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        result += cell.getStringCellValue() + "=";
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
 
                    case Cell.CELL_TYPE_FORMULA:
                        result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        result += "|";
                        break;
                }
            }
            result += "\n";
        }
 
        return result;
    }
 
}
