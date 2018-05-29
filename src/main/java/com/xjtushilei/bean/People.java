package com.xjtushilei.bean;

public class People {
    private String 课 = "";

    private int 当月应收件数 = 0;
    private double 当月应收保费 = 0;

    private int 当月已收件数 = 0;
    private double 当月已收保费 = 0;

    private int 当月未收件数 = 0;
    private double 当月未收保费 = 0;

    private double 当月件数达成 = 0.0000;
    private double 当月保费达成 = 0.0000;

    private int 宽末未收件数 = 0;
    private int 宽一未收件数 = 0;

    private int 总未收件数 = 0;

    /**
     * @param 业务员部门id
     */
    public People(String 业务员部门id) {
        super();
        this.set课(业务员部门id);
    }

    /**
     *
     */
    public People() {
        super();
    }

    public String get课() {
        return 课;
    }

    public void set课(String 业务员部门id) {
        if (业务员部门id.equals("11100067713")) {
            this.课 = "5课";
        } else if (业务员部门id.equals("11100073346")) {
            this.课 = "18课";
        } else if (业务员部门id.equals("11100020425")) {
            this.课 = "11课";
        } else if (业务员部门id.equals("11100104129")) {
            this.课 = "22课";
        } else if (业务员部门id.equals("11100112847")) {
            this.课 = "27课";
        } else if (业务员部门id.equals("11100118816")) {
            this.课 = "28课";
        }

    }

    public int get当月应收件数() {
        return 当月应收件数;
    }

    public void set当月应收件数(int 当月应收件数) {
        this.当月应收件数 = 当月应收件数;
    }

    public double get当月应收保费() {
        return 当月应收保费;
    }

    public void set当月应收保费(double 当月应收保费) {
        this.当月应收保费 = 当月应收保费;
    }

    public int get当月已收件数() {
        return 当月已收件数;
    }

    public void set当月已收件数(int 当月已收件数) {
        this.当月已收件数 = 当月已收件数;
    }

    public double get当月已收保费() {
        return 当月已收保费;
    }

    public void set当月已收保费(double 当月已收保费) {
        this.当月已收保费 = 当月已收保费;
    }

    public int get当月未收件数() {
        return 当月未收件数;
    }

    public void set当月未收件数(int 当月未收件数) {
        this.当月未收件数 = 当月未收件数;
    }

    public double get当月未收保费() {
        return 当月未收保费;
    }

    public void set当月未收保费(double 当月未收保费) {
        this.当月未收保费 = 当月未收保费;
    }

    public int get宽末未收件数() {
        return 宽末未收件数;
    }

    public void set宽末未收件数(int 宽末未收件数) {
        this.宽末未收件数 = 宽末未收件数;
    }

    public int get宽一未收件数() {
        return 宽一未收件数;
    }

    public void set宽一未收件数(int 宽一未收件数) {
        this.宽一未收件数 = 宽一未收件数;
    }

    public int get总未收件数() {
        return 总未收件数;
    }

    public void set总未收件数(int 总未收件数) {
        this.总未收件数 = 总未收件数;
    }

    public double get当月件数达成() {
        return 当月件数达成;
    }

    public void set当月件数达成(double 当月件数达成) {
        this.当月件数达成 = 当月件数达成;
    }

    public double get当月保费达成() {
        return 当月保费达成;
    }

    public void set当月保费达成(double 当月保费达成) {
        this.当月保费达成 = 当月保费达成;
    }

}
