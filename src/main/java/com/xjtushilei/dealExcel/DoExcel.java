package com.xjtushilei.dealExcel;

import com.xjtushilei.bean.People;
import jxl.CellView;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.biff.DisplayFormat;
import jxl.format.Alignment;
import jxl.format.BorderLineStyle;
import jxl.format.VerticalAlignment;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.Number;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.*;
import java.util.Map.Entry;

public class DoExcel {

    private final String path当月已收 = "D:/报表/放报表/当月已收.xls";
    private final String path当月应收 = "D:/报表/放报表/当月应收.xls";
    private final String path宽末未收 = "D:/报表/放报表/宽末未收.xls";
    private final String path宽一未收 = "D:/报表/放报表/宽一未收.xls";
    private String path生成文件 = "D:/报表/生成报表/生成文件"; // 后面生成带日期时间的文件名
    private LinkedHashMap<String, People> allPeople = null;

    public static void main(String[] args) {

    }

    public static String 距离80的函数(int 当月应收件数, int 当月已收件数) {
        double temp = 当月应收件数 * 0.8 - 当月已收件数;
        int result = (int) Math.ceil(temp);
        String sresult = "";
        if (result <= 0) {
            sresult = " ";
        } else {
            sresult = result + "";
        }
        return sresult;

    }

    public static int getMounth(String what) {
        switch (what) {
            case "当":
                return LocalDate.now().getMonthValue();
            case "宽末":
                int now = LocalDate.now().getMonthValue();

                if (now == 1) {
                    return 11;
                } else if (now == 2) {
                    return 12;
                } else {
                    return now - 2;
                }

            case "宽一":
                int now1 = LocalDate.now().getMonthValue();
                if (now1 == 1) {
                    return 12;
                } else {
                    return now1 - 1;
                }
            default:
                return LocalDate.now().getMonthValue();
        }

    }

    public void writeStyle() {

        try {
            Thread.sleep(1000);
        } catch (InterruptedException e2) {
            e2.printStackTrace();
        }
        System.out.println("\n开始更改样式！\n");
        jxl.write.WritableCellFormat 大标题样式 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 17, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            大标题样式 = new jxl.write.WritableCellFormat(wf);
            大标题样式.setBackground(jxl.format.Colour.WHITE);
            大标题样式.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            大标题样式.setAlignment(Alignment.CENTRE);
            大标题样式.setWrap(true);
            大标题样式.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 标题样式 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            标题样式 = new jxl.write.WritableCellFormat(wf);
            标题样式.setBackground(jxl.format.Colour.WHITE);
            标题样式.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            标题样式.setAlignment(Alignment.CENTRE);
            标题样式.setWrap(true);
            标题样式.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }
        DisplayFormat DisplayFormat = NumberFormats.PERCENT_INTEGER;
        jxl.write.WritableCellFormat 件数达成样式 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.RED);
            件数达成样式 = new jxl.write.WritableCellFormat(wf, DisplayFormat);
            件数达成样式.setBackground(jxl.format.Colour.PALE_BLUE);
            件数达成样式.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            件数达成样式.setAlignment(Alignment.CENTRE);
            件数达成样式.setVerticalAlignment(VerticalAlignment.CENTRE);
            件数达成样式.setWrap(true);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 保费达成样式 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.RED);
            保费达成样式 = new jxl.write.WritableCellFormat(wf, DisplayFormat);
            保费达成样式.setBackground(jxl.format.Colour.WHITE);
            保费达成样式.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            保费达成样式.setAlignment(Alignment.CENTRE);
            保费达成样式.setWrap(true);
            保费达成样式.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 总未收件数 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.RED);
            总未收件数 = new jxl.write.WritableCellFormat(wf);
            总未收件数.setBackground(jxl.format.Colour.WHITE);
            总未收件数.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            总未收件数.setAlignment(Alignment.CENTRE);
            总未收件数.setWrap(true);
            总未收件数.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 非零的样式 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.RED);
            非零的样式 = new jxl.write.WritableCellFormat(wf);
            非零的样式.setBackground(jxl.format.Colour.YELLOW);
            非零的样式.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            非零的样式.setAlignment(Alignment.CENTRE);
            非零的样式.setWrap(true);
            非零的样式.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 祝贺达成 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            祝贺达成 = new jxl.write.WritableCellFormat(wf);
            祝贺达成.setBackground(jxl.format.Colour.LIME);
            祝贺达成.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            祝贺达成.setAlignment(Alignment.CENTRE);
            祝贺达成.setWrap(true);
            祝贺达成.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 一步之遥 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            一步之遥 = new jxl.write.WritableCellFormat(wf);
            一步之遥.setBackground(jxl.format.Colour.YELLOW);
            一步之遥.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            一步之遥.setAlignment(Alignment.CENTRE);
            一步之遥.setWrap(true);
            一步之遥.setVerticalAlignment(VerticalAlignment.CENTRE);

        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 加油 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            加油 = new jxl.write.WritableCellFormat(wf);
            加油.setBackground(jxl.format.Colour.ORANGE);
            加油.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            加油.setAlignment(Alignment.CENTRE);
            加油.setWrap(true);
            加油.setVerticalAlignment(VerticalAlignment.CENTRE);
        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        jxl.write.WritableCellFormat 需改善 = null;
        try {
            jxl.write.WritableFont wf = new jxl.write.WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
            wf.setColour(jxl.format.Colour.BLACK);
            需改善 = new jxl.write.WritableCellFormat(wf);
            需改善.setBackground(jxl.format.Colour.TAN);
            需改善.setBorder(jxl.format.Border.ALL, BorderLineStyle.THIN);
            需改善.setAlignment(Alignment.CENTRE);
            需改善.setWrap(true);
            需改善.setVerticalAlignment(VerticalAlignment.CENTRE);
        } catch (WriteException e1) {
            e1.printStackTrace();
        }

        try {
            Workbook rwb = Workbook.getWorkbook(new File(path生成文件));// 原xls文件
            WritableWorkbook wwb = Workbook.createWorkbook(new File(path生成文件), rwb);// 临时xls文件
            for (int count_sheet = 0; count_sheet <= 4; count_sheet++) {
                WritableSheet sheet = wwb.getSheet(count_sheet);// 工作表
                sheet.getWritableCell(0, 0).setCellFormat(大标题样式);

                for (int j = 1; j < sheet.getRows(); j++) {
                    for (int i = 0; i <= 13; i++) {
                        if (i >= 0 && i <= 6) {
                            sheet.getWritableCell(i, j).setCellFormat(标题样式);
                        }
                        if (i >= 9 && i <= 10) {
                            if (j == 1) {
                                sheet.getWritableCell(i, j).setCellFormat(标题样式);

                            } else {
                                if (!sheet.getCell(i, j).getContents().equals("0")) {
                                    sheet.getWritableCell(i, j).setCellFormat(非零的样式);
                                } else {
                                    sheet.getWritableCell(i, j).setCellFormat(标题样式);
                                }

                            }
                        }
                    }
                    sheet.getWritableCell(7, j).setCellFormat(件数达成样式);
                    sheet.getWritableCell(8, j).setCellFormat(保费达成样式);
                    sheet.getWritableCell(11, j).setCellFormat(总未收件数);
                    sheet.getWritableCell(12, j).setCellFormat(标题样式);
                }

                for (int i = 0; i <= 11; i++) {
                    sheet.setColumnView(i, 10);
                }
                sheet.setColumnView(12, 15);
                sheet.setRowView(0, 800);
                sheet.setRowView(1, 800);
            }

            wwb.write();
            wwb.close();
            rwb.close();
            System.out.println("\n样式更改完成 !\n");
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        System.out.println("最终文件 :  <" + path生成文件 + "> \n");
        System.out.println("祝彤彤妈妈工作顺利！生活愉快！");
    }

    public void init() {
        allPeople = new LinkedHashMap<String, People>();

        SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月dd日HH时mm分ss秒");// 设置日期格式
        path生成文件 = path生成文件 + "_" + df.format(new Date()) + ".xls";

    }

    public void do当月已收() {

        Workbook wb = null;
        try {
            wb = Workbook.getWorkbook(new File(path当月已收));
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = wb.getSheet(0); // get sheet(0)
        // 遍历
        for (int i = 1; i < sheet.getRows(); i++) {
            String name = sheet.getCell(5, i).getContents();
            String 业务员部门id = sheet.getCell(7, i).getContents();
            double 当月已收保费 = ((NumberCell) sheet.getCell(4, i)).getValue();

            // System.out.println(name+当月已收保费);

            if (allPeople.containsKey(name)) {
                allPeople.get(name).set当月已收件数(allPeople.get(name).get当月已收件数() + 1);
                allPeople.get(name).set当月已收保费(allPeople.get(name).get当月已收保费() + 当月已收保费);
            } else {
                allPeople.put(name, new People(业务员部门id));
                allPeople.get(name).set当月已收件数(allPeople.get(name).get当月已收件数() + 1);
                allPeople.get(name).set当月已收保费(allPeople.get(name).get当月已收保费() + 当月已收保费);
            }
        }
        System.out.println("处理" + "  <" + path当月已收 + ">  完毕!\n");
    }

    public void do当月应收() {
        // HashMap<String, People> allPeople=new HashMap<>();
        Workbook wb = null;
        try {
            wb = Workbook.getWorkbook(new File(path当月应收));
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = wb.getSheet(0); // get sheet(0)
        // 遍历
        for (int i = 1; i < sheet.getRows(); i++) {
            String name = sheet.getCell(9, i).getContents();
            String 业务员部门id = sheet.getCell(12, i).getContents();
            double 当月应收保费 = ((NumberCell) sheet.getCell(5, i)).getValue();

            // System.out.println(name+当月应收保费);

            if (allPeople.containsKey(name)) {
                allPeople.get(name).set当月应收件数(allPeople.get(name).get当月应收件数() + 1);
                allPeople.get(name).set当月应收保费(allPeople.get(name).get当月应收保费() + 当月应收保费);
            } else {
                allPeople.put(name, new People(业务员部门id));
                allPeople.get(name).set当月应收件数(allPeople.get(name).get当月应收件数() + 1);
                allPeople.get(name).set当月应收保费(allPeople.get(name).get当月应收保费() + 当月应收保费);
            }
        }
        System.out.println("处理" + "  <" + path当月应收 + ">  完毕!\n");
    }

    public void do宽末未收() {
        Workbook wb = null;
        try {
            wb = Workbook.getWorkbook(new File(path宽末未收));
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = wb.getSheet(0); // get sheet(0)
        // 遍历
        for (int i = 1; i < sheet.getRows(); i++) {
            String name = sheet.getCell(9, i).getContents();
            String 业务员部门id = sheet.getCell(12, i).getContents();

            // System.out.println(name);

            if (allPeople.containsKey(name)) {
                allPeople.get(name).set宽末未收件数(allPeople.get(name).get宽末未收件数() + 1);
            } else {
                allPeople.put(name, new People(业务员部门id));
                allPeople.get(name).set宽末未收件数(allPeople.get(name).get宽末未收件数() + 1);
            }
        }
        System.out.println("处理" + "  <" + path宽末未收 + ">  完毕!\n");
    }

    public void do宽一未收() {
        Workbook wb = null;
        try {
            wb = Workbook.getWorkbook(new File(path宽一未收));
        } catch (BiffException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = wb.getSheet(0); // get sheet(0)
        // 遍历
        for (int i = 1; i < sheet.getRows(); i++) {
            String name = sheet.getCell(9, i).getContents();
            String 业务员部门id = sheet.getCell(12, i).getContents();

            // System.out.println(name);

            if (allPeople.containsKey(name)) {
                allPeople.get(name).set宽一未收件数(allPeople.get(name).get宽一未收件数() + 1);
            } else {
                allPeople.put(name, new People(业务员部门id));
                allPeople.get(name).set宽一未收件数(allPeople.get(name).get宽一未收件数() + 1);
            }
        }
        System.out.println("处理" + "  <" + path宽一未收 + ">  完毕!\n");
    }

    public void do计算() {
        // 依次计算所有的人的数据
        for (String key : allPeople.keySet()) {
            String name = key;

            int 当月未收件数 = allPeople.get(name).get当月应收件数() - allPeople.get(name).get当月已收件数();
            allPeople.get(name).set当月未收件数(当月未收件数);

            double 当月未收保费 = allPeople.get(name).get当月应收保费() - allPeople.get(name).get当月已收保费();
            allPeople.get(name).set当月未收保费(当月未收保费);

            int 总未收件数 = allPeople.get(name).get当月未收件数() + allPeople.get(name).get宽末未收件数()
                    + allPeople.get(name).get宽一未收件数();
            allPeople.get(name).set总未收件数(总未收件数);

            if (!(Double.compare(allPeople.get(name).get当月应收件数(), 0) == 0)) {
                double 当月件数达成 = (double) allPeople.get(name).get当月已收件数() / (double) allPeople.get(name).get当月应收件数();
                allPeople.get(name).set当月件数达成(当月件数达成);
            }
            if (!(Double.compare(allPeople.get(name).get当月应收保费(), 0) == 0)) {
                double 当月保费达成 = allPeople.get(name).get当月已收保费() / allPeople.get(name).get当月应收保费();
                allPeople.get(name).set当月保费达成(当月保费达成);
            }
        }
        System.out.println("计算完毕！\n");
    }

    public void do排序() {
        List<Map.Entry<String, People>> infoIds = new ArrayList<Map.Entry<String, People>>(allPeople.entrySet());

        Collections.sort(infoIds, new Comparator<Map.Entry<String, People>>() {
            public int compare(Map.Entry<String, People> o1, Map.Entry<String, People> o2) {
                People p1 = (People) o1.getValue();
                People p2 = (People) o2.getValue();
                ;

                double 当月件数达成1 = p1.get当月件数达成();
                double 当月件数达成2 = p2.get当月件数达成();
                return Double.compare(当月件数达成2, 当月件数达成1);
            }
        });
        /* 转换成新map输出 */
        LinkedHashMap<String, People> newMap = new LinkedHashMap<String, People>();

        for (Map.Entry<String, People> entity : infoIds) {
            newMap.put(entity.getKey(), entity.getValue());
        }
        allPeople = newMap;
        System.out.println("排序完毕！\n");
    }

    public void 写excel() {
        System.out.println("开始生成文件！\n");

        // 创建一个可写入的工作表
        WritableWorkbook wwb = null;
        try {
            wwb = Workbook.createWorkbook(new File(path生成文件));
        } catch (IOException e) {
            e.printStackTrace();
        }

        WritableSheet sheet5课 = wwb.createSheet("5课", 0);
        WritableSheet sheet11课 = wwb.createSheet("11课", 1);
        WritableSheet sheet18课 = wwb.createSheet("18课", 2);
        WritableSheet sheet22课 = wwb.createSheet("22课", 3);
        WritableSheet sheet27课 = wwb.createSheet("27课", 4);
        WritableSheet sheet28课 = wwb.createSheet("28课", 5);
        String[] 列标题 = {"业务员", getMounth("当") + "月应\n收件数", getMounth("当") + "月应\n收保费", getMounth("当") + "月已\n收件数",
                getMounth("当") + "月已\n收保费", getMounth("当") + "月未\n收件数", getMounth("当") + "月未\n收保费",
                getMounth("当") + "月件\n数达成", getMounth("当") + "月保\n费达成", getMounth("宽末") + "月未\n收件数",
                getMounth("宽一") + "月未\n收件数", "总未收\n件数", getMounth("当") + "月距 80% \n差额件数"};
        try {
            sheet5课.mergeCells(0, 0, 12, 0);
            sheet11课.mergeCells(0, 0, 12, 0);
            sheet18课.mergeCells(0, 0, 12, 0);
            sheet22课.mergeCells(0, 0, 12, 0);
            sheet27课.mergeCells(0, 0, 12, 0);
            sheet28课.mergeCells(0, 0, 12, 0);
            for (int i = 0; i < 列标题.length; i++) {
                // 设置列宽
                CellView cellView = new CellView();
                cellView.setSize(3000);
                sheet5课.setColumnView(i, cellView);
                sheet11课.setColumnView(i, cellView);
                sheet18课.setColumnView(i, cellView);
                sheet22课.setColumnView(i, cellView);
                sheet27课.setColumnView(i, cellView);
                sheet28课.setColumnView(i, cellView);

                sheet5课.addCell(new Label(i, 1, 列标题[i]));
                sheet11课.addCell(new Label(i, 1, 列标题[i]));
                sheet18课.addCell(new Label(i, 1, 列标题[i]));
                sheet22课.addCell(new Label(i, 1, 列标题[i]));
                sheet27课.addCell(new Label(i, 1, 列标题[i]));
                sheet28课.addCell(new Label(i, 1, 列标题[i]));

            }
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }

        int 总已收件数5课 = 0;
        int 总已收件数11课 = 0;
        int 总已收件数18课 = 0;
        int 总已收件数22课 = 0;
        int 总已收件数27课 = 0;
        int 总已收件数28课 = 0;

        int 总应该收件数5课 = 0;
        int 总应该收件数11课 = 0;
        int 总应该收件数18课 = 0;
        int 总应该收件数22课 = 0;
        int 总应该收件数27课 = 0;
        int 总应该收件数28课 = 0;

        String 整体达标5课 = "";
        String 整体达标11课 = "";
        String 整体达标18课 = "";
        String 整体达标22课 = "";
        String 整体达标27课 = "";
        String 整体达标28课 = "";

        // 记录当前写到了第几行
        int count5 = 2;
        int count11 = 2;
        int count18 = 2;
        int count22 = 2;
        int count27 = 2;
        int count28 = 2;

        // 定义百分数形式
        WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false);
        DisplayFormat DisplayFormat = NumberFormats.PERCENT_INTEGER;
        WritableCellFormat wcfF = new WritableCellFormat(wf, DisplayFormat);

        for (Entry<String, People> entry : allPeople.entrySet()) {
            String name = entry.getKey();
            String 课 = entry.getValue().get课();
            int 当月应收件数 = entry.getValue().get当月应收件数();
            double 当月应收保费 = entry.getValue().get当月应收保费();
            int 当月已收件数 = entry.getValue().get当月已收件数();
            double 当月已收保费 = entry.getValue().get当月已收保费();
            int 当月未收件数 = entry.getValue().get当月未收件数();
            double 当月未收保费 = entry.getValue().get当月未收保费();
            if (当月未收保费 < 0) {
                当月未收保费 = 0;
            }
            double 当月件数达成 = entry.getValue().get当月件数达成();
            double 当月保费达成 = entry.getValue().get当月保费达成();
            if (当月保费达成 > 1.0000000001) {
                当月保费达成 = 1;
            }
            int 宽末未收件数 = entry.getValue().get宽末未收件数();
            int 宽一未收件数 = entry.getValue().get宽一未收件数();
            int 总未收件数 = entry.getValue().get总未收件数();

            if (课.equals("5课")) {
                总已收件数5课 = 总已收件数5课 + 当月已收件数;
                总应该收件数5课 = 总应该收件数5课 + 当月应收件数;

                try {
                    sheet5课.addCell(new Label(0, count5, name));
                    sheet5课.addCell(new Number(1, count5, 当月应收件数));
                    sheet5课.addCell(new Label(2, count5, (int) Math.round(当月应收保费) + ""));
                    sheet5课.addCell(new Number(3, count5, 当月已收件数));
                    sheet5课.addCell(new Label(4, count5, (int) Math.round(当月已收保费) + ""));

                    sheet5课.addCell(new Number(5, count5, 当月未收件数));
                    sheet5课.addCell(new Label(6, count5, (int) Math.round(当月未收保费) + ""));

                    sheet5课.addCell(new Number(7, count5, 当月件数达成, wcfF));
                    sheet5课.addCell(new Number(8, count5, 当月保费达成, wcfF));
                    sheet5课.addCell(new Number(9, count5, 宽末未收件数));
                    sheet5课.addCell(new Number(10, count5, 宽一未收件数));
                    sheet5课.addCell(new Number(11, count5, 总未收件数));
                    sheet5课.addCell(new Label(12, count5, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count5++;

            } else if (课.equals("11课")) {
                总已收件数11课 = 总已收件数11课 + 当月已收件数;
                总应该收件数11课 = 总应该收件数11课 + 当月应收件数;

                try {
                    sheet11课.addCell(new Label(0, count11, name));
                    sheet11课.addCell(new Number(1, count11, 当月应收件数));
                    sheet11课.addCell(new Label(2, count11, (int) Math.round(当月应收保费) + ""));

                    sheet11课.addCell(new Number(3, count11, 当月已收件数));
                    sheet11课.addCell(new Label(4, count11, (int) Math.round(当月已收保费) + ""));

                    sheet11课.addCell(new Number(5, count11, 当月未收件数));
                    sheet11课.addCell(new Label(6, count11, (int) Math.round(当月未收保费) + ""));

                    sheet11课.addCell(new Number(7, count11, 当月件数达成, wcfF));
                    sheet11课.addCell(new Number(8, count11, 当月保费达成, wcfF));
                    sheet11课.addCell(new Number(9, count11, 宽末未收件数));
                    sheet11课.addCell(new Number(10, count11, 宽一未收件数));
                    sheet11课.addCell(new Number(11, count11, 总未收件数));
                    sheet11课.addCell(new Label(12, count11, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count11++;

            } else if (课.equals("18课")) {
                总已收件数18课 = 总已收件数18课 + 当月已收件数;
                总应该收件数18课 = 总应该收件数18课 + 当月应收件数;

                try {
                    sheet18课.addCell(new Label(0, count18, name));
                    sheet18课.addCell(new Number(1, count18, 当月应收件数));
                    sheet18课.addCell(new Label(2, count18, (int) Math.round(当月应收保费) + ""));

                    sheet18课.addCell(new Number(3, count18, 当月已收件数));
                    sheet18课.addCell(new Label(4, count18, (int) Math.round(当月已收保费) + ""));

                    sheet18课.addCell(new Number(5, count18, 当月未收件数));
                    sheet18课.addCell(new Label(6, count18, (int) Math.round(当月未收保费) + ""));

                    sheet18课.addCell(new Number(7, count18, 当月件数达成, wcfF));
                    sheet18课.addCell(new Number(8, count18, 当月保费达成, wcfF));
                    sheet18课.addCell(new Number(9, count18, 宽末未收件数));
                    sheet18课.addCell(new Number(10, count18, 宽一未收件数));
                    sheet18课.addCell(new Number(11, count18, 总未收件数));
                    sheet18课.addCell(new Label(12, count18, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count18++;
            } else if (课.equals("22课")) {
                总已收件数22课 = 总已收件数22课 + 当月已收件数;
                总应该收件数22课 = 总应该收件数22课 + 当月应收件数;

                try {
                    sheet22课.addCell(new Label(0, count22, name));
                    sheet22课.addCell(new Number(1, count22, 当月应收件数));
                    sheet22课.addCell(new Label(2, count22, (int) Math.round(当月应收保费) + ""));

                    sheet22课.addCell(new Number(3, count22, 当月已收件数));
                    sheet22课.addCell(new Label(4, count22, (int) Math.round(当月已收保费) + ""));

                    sheet22课.addCell(new Number(5, count22, 当月未收件数));
                    sheet22课.addCell(new Label(6, count22, (int) Math.round(当月未收保费) + ""));

                    sheet22课.addCell(new Number(7, count22, 当月件数达成, wcfF));
                    sheet22课.addCell(new Number(8, count22, 当月保费达成, wcfF));
                    sheet22课.addCell(new Number(9, count22, 宽末未收件数));
                    sheet22课.addCell(new Number(10, count22, 宽一未收件数));
                    sheet22课.addCell(new Number(11, count22, 总未收件数));
                    sheet22课.addCell(new Label(12, count22, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count22++;
            } else if (课.equals("27课")) {
                总已收件数27课 = 总已收件数27课 + 当月已收件数;
                总应该收件数27课 = 总应该收件数27课 + 当月应收件数;

                try {
                    sheet27课.addCell(new Label(0, count27, name));
                    sheet27课.addCell(new Number(1, count27, 当月应收件数));
                    sheet27课.addCell(new Label(2, count27, (int) Math.round(当月应收保费) + ""));

                    sheet27课.addCell(new Number(3, count27, 当月已收件数));
                    sheet27课.addCell(new Label(4, count27, (int) Math.round(当月已收保费) + ""));

                    sheet27课.addCell(new Number(5, count27, 当月未收件数));
                    sheet27课.addCell(new Label(6, count27, (int) Math.round(当月未收保费) + ""));

                    sheet27课.addCell(new Number(7, count27, 当月件数达成, wcfF));
                    sheet27课.addCell(new Number(8, count27, 当月保费达成, wcfF));
                    sheet27课.addCell(new Number(9, count27, 宽末未收件数));
                    sheet27课.addCell(new Number(10, count27, 宽一未收件数));
                    sheet27课.addCell(new Number(11, count27, 总未收件数));
                    sheet27课.addCell(new Label(12, count27, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count27++;
            } else if (课.equals("28课")) {
                总已收件数28课 = 总已收件数28课 + 当月已收件数;
                总应该收件数28课 = 总应该收件数28课 + 当月应收件数;

                try {
                    sheet28课.addCell(new Label(0, count28, name));
                    sheet28课.addCell(new Number(1, count28, 当月应收件数));
                    sheet28课.addCell(new Label(2, count28, (int) Math.round(当月应收保费) + ""));

                    sheet28课.addCell(new Number(3, count28, 当月已收件数));
                    sheet28课.addCell(new Label(4, count28, (int) Math.round(当月已收保费) + ""));

                    sheet28课.addCell(new Number(5, count28, 当月未收件数));
                    sheet28课.addCell(new Label(6, count28, (int) Math.round(当月未收保费) + ""));

                    sheet28课.addCell(new Number(7, count28, 当月件数达成, wcfF));
                    sheet28课.addCell(new Number(8, count28, 当月保费达成, wcfF));
                    sheet28课.addCell(new Number(9, count28, 宽末未收件数));
                    sheet28课.addCell(new Number(10, count28, 宽一未收件数));
                    sheet28课.addCell(new Number(11, count28, 总未收件数));
                    sheet28课.addCell(new Label(12, count28, 距离80的函数(当月应收件数, 当月已收件数)));

                } catch (RowsExceededException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
                count28++;
            }
        }

        整体达标5课 = 计算百分比(总已收件数5课, 总应该收件数5课, 2);
        整体达标11课 = 计算百分比(总已收件数11课, 总应该收件数11课, 2);
        整体达标18课 = 计算百分比(总已收件数18课, 总应该收件数18课, 2);
        整体达标22课 = 计算百分比(总已收件数22课, 总应该收件数22课, 2);
        整体达标27课 = 计算百分比(总已收件数27课, 总应该收件数27课, 2);
        整体达标28课 = 计算百分比(总已收件数28课, 总应该收件数28课, 2);

        try {
            sheet5课.addCell(new Label(0, 0, "5课" + getMounth("当") + "月件数整体达成" + 整体达标5课 + "%"));
            sheet11课.addCell(new Label(0, 0, "11课" + getMounth("当") + "月件数整体达成" + 整体达标11课 + "%"));
            sheet18课.addCell(new Label(0, 0, "18课" + getMounth("当") + "月件数整体达成" + 整体达标18课 + "%"));
            sheet22课.addCell(new Label(0, 0, "22课" + getMounth("当") + "月件数整体达成" + 整体达标22课 + "%"));
            sheet27课.addCell(new Label(0, 0, "27课" + getMounth("当") + "月件数整体达成" + 整体达标27课 + "%"));
            sheet28课.addCell(new Label(0, 0, "28课" + getMounth("当") + "月件数整体达成" + 整体达标28课 + "%"));

        } catch (RowsExceededException e1) {
            e1.printStackTrace();
        } catch (WriteException e1) {
            e1.printStackTrace();
        }
        try {
            // 从内存中写入文件中
            wwb.write();
            // 关闭资源，释放内存
            wwb.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }

        System.out.println("生成excel!");

    }

    public String 计算百分比(double 已收, double 应收, int 保留位数) {
        String result;
        NumberFormat numberFormat = NumberFormat.getInstance();
        // 设置精确到小数点后2位
        numberFormat.setMaximumFractionDigits(2);
        result = numberFormat.format((float) 已收 / (float) 应收 * 100);
        return result;
    }

    public LinkedHashMap<String, People> getAllPeople() {
        return allPeople;
    }

    public void setAllPeople(LinkedHashMap<String, People> allPeople) {
        this.allPeople = allPeople;
    }

}
