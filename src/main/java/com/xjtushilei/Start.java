package com.xjtushilei;

import com.xjtushilei.dealExcel.DoExcel;

public class Start {

    public static void main(String[] args) {

        DoExcel 彤彤妈妈的小机器人 = new DoExcel();

        彤彤妈妈的小机器人.addMap("11100067713", "5课");
        彤彤妈妈的小机器人.addMap("11100020425", "11课");
        彤彤妈妈的小机器人.addMap("11100073346", "18课");
        彤彤妈妈的小机器人.addMap("11100104129", "22课");
        彤彤妈妈的小机器人.addMap("11100112847", "27课");
        彤彤妈妈的小机器人.addMap("11100118816", "28课");

        彤彤妈妈的小机器人.init();
        彤彤妈妈的小机器人.do当月已收();
        彤彤妈妈的小机器人.do当月应收();
        彤彤妈妈的小机器人.do宽末未收();
        彤彤妈妈的小机器人.do宽一未收();
        彤彤妈妈的小机器人.do计算();
        彤彤妈妈的小机器人.do排序();
        彤彤妈妈的小机器人.new写excel();
        彤彤妈妈的小机器人.writeStyle();

    }

}
