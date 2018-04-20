/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.excelocenka;

import com.tender.entity.Bal;
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
import java.util.Collection;
import java.util.Collections;
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

            while (objT.get(pos).getLot() == (i + 1) && pos <= objT.size() - 2) {

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
                //     System.out.println(pos);
            }

            znach = new Znach((int) (i + 1), maxC, minC, maxO, minO);
            znachs.add(znach);

        }

        for (int i = 0; i < objT.size(); i++) {

            int maxO = znachs.get(objT.get(i).getLot() - 1).getOtsMax();
            int minO = znachs.get(objT.get(i).getLot() - 1).getOtsMin();
            float maxC = znachs.get(objT.get(i).getLot() - 1).getCenaMax();
            float minC = znachs.get(objT.get(i).getLot() - 1).getCenaMin();

            float cenaK = (float) 0.8;
            float otsK = (float) 0.2;

            //=1+(МАКС($E$5:$E$8)-E5)/(МАКС($E$5:$E$8)-МИН($E$5:$E$8))*9
            if (maxC != minC) {
                objT.get(i).setBalC(1 + (maxC - objT.get(i).getCenO()) / (maxC - minC) * 9);
            } else {
                objT.get(i).setBalC((float) 1.0);
            }
            objT.get(i).setBalCk(objT.get(i).getBalC() * cenaK);

            //=1+(F5-МИН($F$5:F$8))/(МАКС($F$5:F$8)-МИН($F$5:F$8))*9
            if (maxO != minO) {
                objT.get(i).setBalO(1 + (objT.get(i).getOts() - minO) / (maxO - minO) * 9);
            } else {
                objT.get(i).setBalO((float) 1.0);
            }
            objT.get(i).setBalOk(objT.get(i).getBalO() * otsK);

            objT.get(i).setBalOb(objT.get(i).getBalOk() + objT.get(i).getBalCk());
        }

        ArrayList<Bal> bals = new ArrayList<Bal>();
        for (int i = 0; i < objT.size(); i++) {
            Bal bal = new Bal();
            bal.setPos(i);
            bal.setLot(objT.get(i).getLot());
            bal.setBalO(objT.get(i).getBalOb());
            bals.add(bal);
        }
        //int pos =0;
        for (int i = 0; i < 10; i++) {//bals.size()
            System.out.println(bals.get(i).toString());
        }
        
        pos = 0;
        for (int i = 1; i < bals.get(bals.size() - 1).getLot(); i++) {
            int posN = pos;

            while (bals.get(pos).getLot() == i && pos != bals.size()) {
                pos++;
            }

            for (int a = posN+1; a < pos; a++) {
                for (int b = posN; b < pos - a; b++) {
                    if (bals.get(b).getBalO() < bals.get(b + 1).getBalO()) {

                        Bal bufB = new Bal();
                        bufB.setPos(bals.get(b).getPos());
                        bufB.setLot(bals.get(b).getLot());
                        bufB.setBalO(bals.get(b).getBalO());
                        
                        bals.get(b).setPos(bals.get(b + 1).getPos());
                        bals.get(b).setLot(bals.get(b + 1).getLot());
                        bals.get(b).setBalO(bals.get(b + 1).getBalO());
                        
                        bals.get(b + 1).setPos(bufB.getPos());
                        bals.get(b + 1).setLot(bufB.getLot());
                        bals.get(b + 1).setBalO(bufB.getBalO());
                        
                        /*bals.set(b, bals.get(b + 1));
                        bals.set((b + 1), bufB);*/
                    }
                }
            }
            int r = 1;
            for (int z = posN; z < pos; z++) {
                bals.get(z).setRang(r);
                r++;
            }
        }

        //проверка принципа
        System.out.println("До сортировки:");
        
        ArrayList<Double> mas = new ArrayList<Double>();
        
        for (int i = 0; i < 10; i++) {
            mas.add(new Double(((i + 1)*(0.123456789*i))+i));
            System.out.print(mas.get(i) + " ");
        }

        for (int i = 1; i < mas.size(); i++) {
            for (int j = 0; j < mas.size() - i; j++) {
                if (mas.get(j) < mas.get(j + 1)) {

                   Double bufM = new Double(mas.get(j));

                    mas.set(j, mas.get(j + 1));
                    mas.set((j + 1), bufM);
                }
            }
        }
        System.out.println("\nПосле сортировки:");
        
        for (int i = 0; i < mas.size(); i++) {
            System.out.print(mas.get(i) + " ");
        }


        /*for (int i = 0; i < objT.size(); i++) {
            System.out.println(objT.get(i).toString());
        }

        for (int i = 0; i < znachs.size(); i++) {
            System.out.println(znachs.get(i).toString());
        }*/

        for (int i = 0; i < 10; i++) {//bals.size()
            System.out.println(bals.get(i).toString());
        }
        try {
            fis.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
