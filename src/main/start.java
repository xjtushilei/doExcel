package main;

import dealExcel.doExcel;

public class start {

	public static void main(String[] args) {
		
		doExcel 彤彤妈妈的小机器人 = new doExcel();

		彤彤妈妈的小机器人.init();
		彤彤妈妈的小机器人.do当月已收();
		彤彤妈妈的小机器人.do当月应收();
		彤彤妈妈的小机器人.do宽末未收();
		彤彤妈妈的小机器人.do宽一未收();
		彤彤妈妈的小机器人.do计算();
		彤彤妈妈的小机器人.do排序();
		彤彤妈妈的小机器人.写excel();

	}

}
