/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.tender.entity;

/**
 *
 * @author User
 */
public class Znach {
    private int nomL;
    private float cenMax;
    private float cenaMin;
    private int otsMax;
    private int otsMin;

    public Znach() {
    }

    public Znach(int nomL, float cenMax, float cenaMin, int otsMax, int otsMin) {
        this.nomL = nomL;
        this.cenMax = cenMax;
        this.cenaMin = cenaMin;
        this.otsMax = (int) otsMax;
        this.otsMin = (int) otsMin;
    }

    public int getNomL() {
        return nomL;
    }

    public void setNomL(int nomL) {
        this.nomL = nomL;
    }

    public float getCenMax() {
        return cenMax;
    }

    public void setCenMax(float cenMax) {
        this.cenMax = cenMax;
    }

    public float getCenaMin() {
        return cenaMin;
    }

    public void setCenaMin(float cenaMin) {
        this.cenaMin = cenaMin;
    }

    public int getOtsMax() {
        return otsMax;
    }

    public void setOtsMax(int otsMax) {
        this.otsMax = otsMax;
    }

    public int getOtsMin() {
        return otsMin;
    }

    public void setOtsMin(int otsMin) {
        this.otsMin = otsMin;
    }

    @Override
    public String toString() {
        return "Znach{" + "nomL=" + nomL + ", cenMax=" + cenMax + ", cenaMin=" + cenaMin + ", otsMax=" + otsMax + ", otsMin=" + otsMin + '}';
    }

}
