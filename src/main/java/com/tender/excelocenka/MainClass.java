/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.excelocenka;

import java.io.FileNotFoundException;

/**
 *
 * @author User
 */
public class MainClass {
    public static void main(String[] args) throws FileNotFoundException{
        System.out.println(ExcelParser.parse());
    }
}
