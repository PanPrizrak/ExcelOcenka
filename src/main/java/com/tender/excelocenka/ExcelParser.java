/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.excelocenka;

import com.tender.entity.ObjT;
import com.tender.entity.Znach;
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
import java.util.ArrayList;
import java.util.Iterator;
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
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            fis = new FileInputStream(window.getSelectedFile());
        }
        JOptionPane.showMessageDialog(null, window.getSelectedFile().toString());
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
        XSSFSheet spreadsheet = workbook.getSheetAt(3);
        Iterator< Row> rowIterator = spreadsheet.iterator();
        ArrayList<ObjT> objT = new ArrayList<>();
        ObjT buf;
        while (rowIterator.hasNext()) {
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
            buf = new ObjT();
            for (int i = 0; i < 8; i++) {
                switch (i) {
                    case 0:
                        buf.setLot((int) row.getCell(i).getNumericCellValue());
                        break;
                    case 1:
                        if (row.getCell(i) != null) {
                            buf.setNameC(row.getCell(i).getStringCellValue());
                        } else {
                            buf.setNameC("");
                        }
                        break;
                    case 2:
                        if (row.getCell(i) != null) {
                            buf.setOts((int) row.getCell(i).getNumericCellValue());
                        } else {
                            buf.setOts((int) 0);
                        }
                        break;
                    case 3:
                        if (row.getCell(i) != null) {
                            buf.setNameO(row.getCell(i).getStringCellValue());
                        } else {
                            buf.setNameO("");
                        }
                        break;
                    case 4:
                        if (row.getCell(i) != null) {
                            buf.setEd(row.getCell(i).getStringCellValue());
                        } else {
                            buf.setEd("");
                        }
                        break;
                    case 5:
                        if (row.getCell(i) != null) {
                            buf.setCen((float) row.getCell(i).getNumericCellValue());
                        } else {
                            buf.setCen((float) 0.0);
                        }
                        break;
                    case 6:
                        if (row.getCell(i) != null) {
                            buf.setCenS((float) row.getCell(i).getNumericCellValue());
                        } else {
                            buf.setCenS((float) 0.0);
                        }
                        break;
                    case 7:
                        if (row.getCell(i) != null) {
                            buf.setCenO((float) row.getCell(i).getNumericCellValue());
                        } else {
                            buf.setCenO((float) 0.0);
                        }
                        break;
                }

            }
            //System.out.println(buf.toString());
            objT.add(buf);
        }

        ArrayList<Znach> znachs = new ArrayList<Znach>();
        Znach znach;
        int pos = 0;
        for (int i = 0; i < objT.get(objT.size() - 1).getLot(); i++) {

            //int lot = pos-1;
            float minC = objT.get(pos).getCenO();
            float maxC = objT.get(pos).getCenO();
            int minO = objT.get(pos).getOts();
            int maxO = objT.get(pos).getOts();

            while (objT.get(pos).getLot() == (i + 1) && pos <= objT.size()-2) {

                if (objT.get(pos).getCenO() > maxC) {
                    maxC = objT.get(pos).getCenO();
                }
                if (objT.get(pos).getCenO() < minC) {
                    minC = objT.get(pos).getCenO();
                }

                if (objT.get(pos).getOts() > maxO) {
                    maxO = objT.get(pos).getOts();
                }
                if (objT.get(pos).getOts() < minO) {
                    minO = objT.get(pos).getOts();
                }
                if (pos < objT.size() - 1) {
                    pos++;
                }//if
                System.out.println(pos);
            }

            znach = new Znach((int) (i + 1), maxC, minC, maxO, minO);
            znachs.add(znach);

        }
        
        for (int i = 0; i < objT.size(); i++){
            objT.get(i).setBalC(1+(znachs.get(objT.get(i).getLot()).getCenMax()-objT.get(i).getCenO())/(znachs.get(objT.get(i).getLot()).getCenMax()-znachs.get(objT.get(i).getLot()).getCenaMin())*9);
        }

        for (int i = 0; i < objT.size(); i++) {
            System.out.println(objT.get(i).toString());
        }

        for (int i = 0; i < znachs.size(); i++) {
            System.out.println(znachs.get(i).toString());
        }
        try {
            fis.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
