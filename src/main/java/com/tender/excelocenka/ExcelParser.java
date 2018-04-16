/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.excelocenka;
 
import com.tender.entity.ObjT;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;


import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

 
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
 
public class ExcelParser {
    static XSSFRow row;
    public static void parse() throws FileNotFoundException {
    //инициализируем потоки
       /* String result = "";
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
 
        return result;*/
    
    
      String result = "";
        FileInputStream fis = null;
        JFileChooser window = new JFileChooser();
        int returnValue = window.showOpenDialog(null);
        if(returnValue==JFileChooser.APPROVE_OPTION)
            fis = new FileInputStream(window.getSelectedFile());
            JOptionPane.showMessageDialog(null, window.getSelectedFile().toString());
      XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
      XSSFSheet spreadsheet = workbook.getSheetAt(3);
      Iterator < Row > rowIterator = spreadsheet.iterator();
      List <ObjT> objT = null;
      while (rowIterator.hasNext()) 
      {
         row = (XSSFRow) rowIterator.next();
         /*Iterator < Cell > cellIterator = row.cellIterator();
         while ( cellIterator.hasNext()) 
         {
            Cell cell = cellIterator.next();
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
            result += "\n";*/
         ObjT buf = null;
         for ( int i = 0 ; i < 8 ; i++)
         {
             buf = new ObjT();
             switch (i) {
                 case 0:
                     buf.setLot((int) row.getCell(i).getNumericCellValue());
                     break;
                 case 1:
                     buf.setNameC(row.getCell(i).getStringCellValue());
                     break;
                 case 2:
                     buf.setOts((int) row.getCell(i).getNumericCellValue());
                     break;
                 case 3:
                     buf.setNameO(row.getCell(i).getStringCellValue());
                     break;
                 case 4:
                     buf.setEd(row.getCell(i).getStringCellValue());
                     break;
                 case 5:
                     buf.setCen((float) row.getCell(i).getNumericCellValue());
                     break;
                 case 6:
                     buf.setCenS((float) row.getCell(i).getNumericCellValue());
                     break;
                 case 7:
                     buf.setCenO((float) row.getCell(i).getNumericCellValue());
                     break;
             }
        }
         objT.add(buf);
      }
        try {
            fis.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
   }
 }
