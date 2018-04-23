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
import java.io.FileOutputStream;
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
    private static float cenaK = (float) 0.8;
    private static float otsK = (float) 0.2;

    private static void getRaschet(ObjT t, float maxC, float minC, int maxO, int minO) {

        if (maxC != minC) {
            t.setBalC((float) (1 + (maxC - t.getCenO()) / (maxC - minC) * 9));
        } else {
            t.setBalC((float) 1.0);
        }
        t.setBalCk(t.getBalC() * cenaK);

        //=1+(F5-МИН($F$5:F$8))/(МАКС($F$5:F$8)-МИН($F$5:F$8))*9
        if (maxO != minO) {
            float b1;
            float b2;
            float b3;
            b1 = (float) t.getOts() - minO;
            b2 = (float) maxO - minO;
            b3 = (float) b1 / b2 * 9;
            t.setBalO((float) (1 + b3));
        } else {
            t.setBalO((float) 1.0);
        }
        t.setBalOk((float) t.getBalO() * otsK);

        t.setBalOb((float) t.getBalOk() + t.getBalCk());

        //return t;
    }

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

        //String result = "";
        FileInputStream fis = null;

        JFileChooser window = new JFileChooser();
        int returnValue = window.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            fis = new FileInputStream(window.getSelectedFile());
        }

        //JOptionPane.showMessageDialog(null, window.getSelectedFile().toString());
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }

        //Заполнение из таблицы экселя
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
            objT.add(buf);
        }

        //Определини экстернов значений
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

                if (objT.get(pos).getCenO() > maxC && objT.get(pos).getCenO() != 0) {
                    maxC = objT.get(pos).getCenO();
                }
                if (objT.get(pos).getCenO() < minC && objT.get(pos).getCenO() != 0) {
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

        //Расчет балов
        for (int i = 0; i < objT.size(); i++) {

            int maxO = znachs.get(objT.get(i).getLot() - 1).getOtsMax();
            int minO = znachs.get(objT.get(i).getLot() - 1).getOtsMin();
            float maxC = znachs.get(objT.get(i).getLot() - 1).getCenaMax();
            float minC = znachs.get(objT.get(i).getLot() - 1).getCenaMin();
            if (objT.get(i).getCenO() != 0) {
                getRaschet(objT.get(i), maxC, minC, maxO, minO);
            }
        }

        /*ObjT z1 = new ObjT(); test getRaschet
        z1 = objT.get(20);
        System.out.println("Z " + z1.toString());
        //ObjT zz =new ObjT();
        z1 = getRaschet(z1, znachs.get(z1.getLot()).getCenaMax(), znachs.get(z1.getLot()).getCenaMin(), znachs.get(z1.getLot()).getOtsMax(), znachs.get(z1.getLot()).getOtsMin());
         System.out.println("Z posle " + z1.toString());
         System.out.println("ObjT " + objT.get(20));*/
        //попарное сравнение
        ArrayList<ObjT> parSrav = new ArrayList<ObjT>();
        for (int i = 0; i < objT.size(); i++) {
            ObjT bufO = new ObjT(objT.get(i));
            //int k = i + 1;
            float maxC;
            float minC;
            int maxO;
            int minO;
            if (i < objT.size() - 1 && objT.get(i).getCenO() != 0 ) {
                //while (objT.get(k).getCenO() != 0 && objT.get(i).getLot() == objT.get(k).getLot()) {
                for (int k = (i + 1); objT.get(i).getLot() == objT.get(k).getLot(); k++) {

                    if ( objT.get(k).getCenO() != 0) {

                        ObjT bufOp = new ObjT(objT.get(k));

                        if (objT.get(i).getCenO() > objT.get(k).getCenO()) {
                            maxC = objT.get(i).getCenO();
                            minC = objT.get(k).getCenO();
                        } else {
                            maxC = objT.get(k).getCenO();
                            minC = objT.get(i).getCenO();
                        }

                        if (objT.get(i).getOts() > objT.get(k).getOts()) {
                            maxO = objT.get(i).getOts();
                            minO = objT.get(k).getOts();
                        } else {
                            maxO = objT.get(k).getOts();
                            minO = objT.get(i).getOts();
                        }

                        getRaschet(bufO, maxC, minC, maxO, minO);
                        getRaschet(bufOp, maxC, minC, maxO, minO);

                        if (bufO.getBalOb() > bufOp.getBalOb()) {
                            bufO.setRang((int) 1);
                            bufOp.setRang((int) 2);
                        } else if (bufO.getBalOb() == bufOp.getBalOb()) {
                            bufO.setRang((int) 1);
                            bufOp.setRang((int) 1);
                        } else {
                            bufO.setRang((int) 2);
                            bufOp.setRang((int) 1);
                        }

                        parSrav.add(bufO);
                        parSrav.add(bufOp);
                    }//if
                }//for k
            }//if
        }//for

        //Заполнение объеекта Bal
        ArrayList<Bal> bals = new ArrayList<Bal>();
        for (int i = 0; i < objT.size(); i++) {
            if (objT.get(i).getCenO() != 0) {
                Bal bal = new Bal();
                bal.setPos(i);
                bal.setLot(objT.get(i).getLot());
                bal.setBalO(objT.get(i).getBalOb());
                bals.add(bal);
            }
        }

        //сортировка  в каждом лоте по убыванию общих балов
        pos = 0;
        for (int i = 1;
                i < bals.get(bals.size() - 1).getLot(); i++) {
            int posN = pos;

            while (bals.get(pos).getLot() == i && pos != bals.size()) {
                pos++;
            }

            for (int a = posN + 1; a < pos; a++) {
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

        //присвоение рангов
        for (int i = 0;
                i < bals.size();
                i++) {
            objT.get(bals.get(i).getPos()).setRang((int) bals.get(i).getRang());
        }
        /*
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
        for (int i = 0; i < 10; i++) {
            System.out.println(objT.get(i).toString());
        }
/*System.out.println("Znach!!!!!!!!!!!!!!!!!!!!!!!!!");
        for (int i = 0; i < 9; i++) {
            System.out.println(znachs.get(i).toString());
        }
System.out.println("Bal!!!!!!!!!!!!!!!!!!!!!!!!!");
        for (int i = 40; i < 51; i++) {//bals.size()
            System.out.println(bals.get(i).toString());
        }*/
        //Экспорт в Excel
        XSSFSheet sheet = workbook.createSheet("Оценка общая");
        /*Object[][] datatypes = {
            {"Datatype", "Type", "Size(in bytes)"},
            {"int", "Primitive", 2},
            {"float", "Primitive", 4},
            {"double", "Primitive", 8},
            {"char", "Primitive", 1},
            {"String", "Non-Primitive", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }*/
//      System.out.println(objT.get(50).toString());
        for (int i = 0;
                i < objT.size();
                i++) {
            Row row = sheet.createRow(i);
            ObjT bufO = new ObjT(objT.get(i));
            for (int j = 0; j < 14; j++) { // меншье 13 тк кол-во полей ObjT 14
                Cell cell = row.createCell(j);
                switch (j) {
                    case 0:
                        cell.setCellValue((int) bufO.getLot());
                        break;
                    case 1:
                        cell.setCellValue((String) bufO.getNameC());
                        break;
                    case 2:
                        cell.setCellValue((int) bufO.getOts());
                        break;
                    case 3:
                        cell.setCellValue((String) bufO.getNameO());
                        break;
                    case 4:
                        cell.setCellValue((String) bufO.getEd());
                        break;
                    case 5:
                        cell.setCellValue((float) bufO.getCen());
                        break;
                    case 6:
                        cell.setCellValue((float) bufO.getCenS());
                        break;
                    case 7:
                        cell.setCellValue((float) bufO.getCenO());
                        break;
                    case 8:
                        cell.setCellValue((float) bufO.getBalC());
                        break;
                    case 9:
                        cell.setCellValue((float) bufO.getBalCk());
                        break;
                    case 10:
                        cell.setCellValue((float) bufO.getBalO());
                        break;
                    case 11:
                        cell.setCellValue((float) bufO.getBalOk());
                        break;
                    case 12:
                        cell.setCellValue((float) bufO.getBalOb());
                        break;
                    case 13:
                        cell.setCellValue(bufO.getRang());
                        break;
                }//switch
            }
        }

        XSSFSheet sheetP = workbook.createSheet("Оценка попарная");
        for (int i = 0;
                i < parSrav.size();
                i++) {
            Row row = sheetP.createRow(i);
            ObjT bufO = new ObjT(parSrav.get(i));
            for (int j = 0; j < 14; j++) { // меншье 13 тк кол-во полей ObjT 14
                Cell cell = row.createCell(j);
                switch (j) {
                    case 0:
                        cell.setCellValue((int) bufO.getLot());
                        break;
                    case 1:
                        cell.setCellValue((String) bufO.getNameC());
                        break;
                    case 2:
                        cell.setCellValue((int) bufO.getOts());
                        break;
                    case 3:
                        cell.setCellValue((String) bufO.getNameO());
                        break;
                    case 4:
                        cell.setCellValue((String) bufO.getEd());
                        break;
                    case 5:
                        cell.setCellValue((float) bufO.getCen());
                        break;
                    case 6:
                        cell.setCellValue((float) bufO.getCenS());
                        break;
                    case 7:
                        cell.setCellValue((float) bufO.getCenO());
                        break;
                    case 8:
                        cell.setCellValue((float) bufO.getBalC());
                        break;
                    case 9:
                        cell.setCellValue((float) bufO.getBalCk());
                        break;
                    case 10:
                        cell.setCellValue((float) bufO.getBalO());
                        break;
                    case 11:
                        cell.setCellValue((float) bufO.getBalOk());
                        break;
                    case 12:
                        cell.setCellValue((float) bufO.getBalOb());
                        break;
                    case 13:
                        cell.setCellValue(bufO.getRang());
                        break;
                }//switch
            }
        }
        
        XSSFSheet znachM = workbook.createSheet("MinCena");
        for (int i = 0; i < znachs.size(); i++) {
            Row row = znachM.createRow(i);
            Znach bufO = new Znach(znachs.get(i));
            for (int j = 0; j < 5; j++) { // меншье 13 тк кол-во полей ObjT 14
                Cell cell = row.createCell(j);
                switch (j) {
                    case 0:
                        cell.setCellValue((int) bufO.getNomL());
                        break;
                    case 1:
                        cell.setCellValue((float) bufO.getCenaMax());
                        break;
                    case 2:
                        cell.setCellValue((float) bufO.getCenaMin());
                        break;
                    case 3:
                        cell.setCellValue((int) bufO.getOtsMax());
                        break;
                    case 4:
                        cell.setCellValue((int) bufO.getOtsMin());
                        break;
                }//switch
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(window.getSelectedFile());
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done, ObjT size = " + objT.size() + ", Bals size = " + bals.size());

        try {
            fis.close();
        } catch (IOException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
