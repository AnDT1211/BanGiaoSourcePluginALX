/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ec.tool.model;

/**
 *
 * @author thangnq
 */
public class NEILModel {
    private int countN;
    private int countE;
    private int countI;
    private int countL;

    public NEILModel(int n, int e, int i, int l) {
        countN = n;
        countE = e;
        countI = i;
        countL = l;
    }

    
    
    public int getTotal() {
        return countN + countE + countI + countL;
    }
    
    public int getCountN() {
        return countN;
    }

    public void setCountN(int countN) {
        this.countN = countN;
    }

    public int getCountE() {
        return countE;
    }

    public void setCountE(int countE) {
        this.countE = countE;
    }

    public int getCountI() {
        return countI;
    }

    public void setCountI(int countI) {
        this.countI = countI;
    }

    public int getCountL() {
        return countL;
    }

    public void setCountL(int countL) {
        this.countL = countL;
    }

    @Override
    public String toString() {
        return "NEILModel{" + "countN=" + countN + ", countE=" + countE + ", countI=" + countI + ", countL=" + countL + '}';
    }
    
    
}
