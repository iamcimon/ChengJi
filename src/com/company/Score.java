package com.company;
public class Score {
   public  String NO;
   public  String partment;
   public  String name;
   public  String fuZhuRen;
   public  String zhuRen;
   public  String lingDao;
   public  String  lastChengji;
   public  double  chengji;
   public  boolean  isTwo;

    public Score() {
        this.isTwo = true;
    }

    public double getChengji() {
        return chengji;
    }

    public void setChengji(double chengji) {
        this.chengji = chengji;
    }

    public boolean isTwo() {
        return isTwo;
    }

    public void setTwo(boolean two) {
        isTwo = two;
    }

    public String getNO() {
        return NO;
    }

    public void setNO(String NO) {
        this.NO = NO;
    }

    public String getPartment() {
        return partment;
    }

    public void setPartment(String partment) {
        this.partment = partment;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getFuZhuRen() {
        return fuZhuRen;
    }

    public void setFuZhuRen(String fuZhuRen) {
        this.fuZhuRen = fuZhuRen;
    }

    public String getZhuRen() {
        return zhuRen;
    }

    public void setZhuRen(String zhuRen) {
        this.zhuRen = zhuRen;
    }

    public String getLingDao() {
        return lingDao;
    }

    public void setLingDao(String lingDao) {
        this.lingDao = lingDao;
    }

    public String getLastChengji() {
        return lastChengji;
    }

    public void setLastChengji(String lastChengji) {
        this.lastChengji = lastChengji;
    }
}
