package dealExcel;

import java.io.File;
import java.io.IOException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import bean.people;
import jxl.CellView;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.biff.DisplayFormat;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class doExcel {

	private final String path当月已收 = "D:/报表/放报表/当月已收.xls";
	private final String path当月应收 = "D:/报表/放报表/当月应收.xls";
	private final String path宽末未收 = "D:/报表/放报表/宽末未收.xls";
	private final String path宽一未收 = "D:/报表/放报表/宽一未收.xls";
	private final String path生成文件 = "D:/报表/生成报表/生成文件"; // 后面生成带日期时间的文件名
	private  LinkedHashMap<String, people> allPeople = null;

	public void init() {
		allPeople = new  LinkedHashMap<String, people>();
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
				allPeople.put(name, new people(业务员部门id));
				allPeople.get(name).set当月已收件数(allPeople.get(name).get当月已收件数() + 1);
				allPeople.get(name).set当月已收保费(allPeople.get(name).get当月已收保费() + 当月已收保费);
			}
		}
		System.out.println("处理"+"  <"+path当月已收+">  完毕!\n");
	}

	public void do当月应收() {
		// HashMap<String, people> allPeople=new HashMap<>();
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
			String name = sheet.getCell(8, i).getContents();
			String 业务员部门id = sheet.getCell(11, i).getContents();
			double 当月应收保费 = ((NumberCell) sheet.getCell(4, i)).getValue();

			// System.out.println(name+当月应收保费);

			if (allPeople.containsKey(name)) {
				allPeople.get(name).set当月应收件数(allPeople.get(name).get当月应收件数() + 1);
				allPeople.get(name).set当月应收保费(allPeople.get(name).get当月应收保费() + 当月应收保费);
			} else {
				allPeople.put(name, new people(业务员部门id));
				allPeople.get(name).set当月应收件数(allPeople.get(name).get当月应收件数() + 1);
				allPeople.get(name).set当月应收保费(allPeople.get(name).get当月应收保费() + 当月应收保费);
			}
		}
		System.out.println("处理"+"  <"+path当月应收+">  完毕!\n");
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
			String name = sheet.getCell(8, i).getContents();
			String 业务员部门id = sheet.getCell(7, i).getContents();

			// System.out.println(name);

			if (allPeople.containsKey(name)) {
				allPeople.get(name).set宽末未收件数(allPeople.get(name).get宽末未收件数() + 1);
			} else {
				allPeople.put(name, new people(业务员部门id));
				allPeople.get(name).set宽末未收件数(allPeople.get(name).get宽末未收件数() + 1);
			}
		}
		System.out.println("处理"+"  <"+path宽末未收+">  完毕!\n");
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
			String name = sheet.getCell(8, i).getContents();
			String 业务员部门id = sheet.getCell(7, i).getContents();

			// System.out.println(name);

			if (allPeople.containsKey(name)) {
				allPeople.get(name).set宽一未收件数(allPeople.get(name).get宽一未收件数() + 1);
			} else {
				allPeople.put(name, new people(业务员部门id));
				allPeople.get(name).set宽一未收件数(allPeople.get(name).get宽一未收件数() + 1);
			}
		}
		System.out.println("处理"+"  <"+path宽一未收+">  完毕!\n");
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
		List<Map.Entry<String, people>> infoIds =new ArrayList<Map.Entry<String, people>>(allPeople.entrySet());
		

		Collections.sort(infoIds, new Comparator<Map.Entry<String, people>>() {   
            public int compare(Map.Entry<String, people> o1, Map.Entry<String, people> o2) {      
            	people p1 = (people) o1.getValue();
            	people p2 = (people) o2.getValue();;
                
            	double 当月件数达成1=p1.get当月件数达成();
            	double 当月件数达成2=p2.get当月件数达成();
                return Double.compare(当月件数达成2, 当月件数达成1);
            }
        });
		/*转换成新map输出*/
        LinkedHashMap<String, people> newMap = new LinkedHashMap <String, people>();
         
        for(Map.Entry<String,people> entity : infoIds){
            newMap.put(entity.getKey(), entity.getValue());
        }
        allPeople=newMap;
		System.out.println("排序完毕！\n");
	}

	public void 写excel() {
		System.out.println("开始生成文件！\n");
		SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月dd日HH时mm分ss秒");// 设置日期格式
		String filepath = path生成文件 + "_" + df.format(new Date()) + ".xls";
		// 创建一个可写入的工作表
		WritableWorkbook wwb = null;
		try {
			wwb = Workbook.createWorkbook(new File(filepath));
		} catch (IOException e) {
			e.printStackTrace();
		}

		WritableSheet sheet5课 = wwb.createSheet("5课", 0);
		WritableSheet sheet11课 = wwb.createSheet("11课", 1);
		WritableSheet sheet18课 = wwb.createSheet("18课", 2);
		WritableSheet sheet22课 = wwb.createSheet("22课", 3);
		String[] 列标题 = { "业务员", "当月应收件数", "当月应收保费", "当月已收件数", "当月已收保费", "当月未收件数", "当月未收保费", "当月件数达成", "当月保费达成",
				"宽末未收件数", "宽一未收件数", "总未收件数" };
		try {
			sheet5课.mergeCells(0, 0, 11, 0);
			sheet11课.mergeCells(0, 0, 11, 0);
			sheet18课.mergeCells(0, 0, 11, 0);
			sheet22课.mergeCells(0, 0, 11, 0);

			for (int i = 0; i < 列标题.length; i++) {
				// 设置列宽
				CellView cellView = new CellView();
				cellView.setSize(3000);
				sheet5课.setColumnView(i, cellView);
				sheet11课.setColumnView(i, cellView);
				sheet18课.setColumnView(i, cellView);
				sheet22课.setColumnView(i, cellView);

				sheet5课.addCell(new Label(i, 1, 列标题[i]));
				sheet11课.addCell(new Label(i, 1, 列标题[i]));
				sheet18课.addCell(new Label(i, 1, 列标题[i]));
				sheet22课.addCell(new Label(i, 1, 列标题[i]));
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

		int 总应该收件数5课 = 0;
		int 总应该收件数11课 = 0;
		int 总应该收件数18课 = 0;
		int 总应该收件数22课 = 0;

		String 整体达标5课 = "";
		String 整体达标11课 = "";
		String 整体达标18课 = "";
		String 整体达标22课 = "";

		// 记录当前写到了第几行
		int count5 = 2;
		int count11 = 2;
		int count18 = 2;
		int count22 = 2;

		// 定义百分数形式
		WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false);
		DisplayFormat DisplayFormat = NumberFormats.PERCENT_INTEGER;
		WritableCellFormat wcfF = new WritableCellFormat(wf, DisplayFormat);

		for (Entry<String, people> entry : allPeople.entrySet()) {
			String name = entry.getKey();
			String 课 = entry.getValue().get课();
			int 当月应收件数 = entry.getValue().get当月应收件数();
			double 当月应收保费 = entry.getValue().get当月应收保费();
			int 当月已收件数 = entry.getValue().get当月已收件数();
			double 当月已收保费 = entry.getValue().get当月已收保费();
			int 当月未收件数 = entry.getValue().get当月未收件数();
			double 当月未收保费 = entry.getValue().get当月未收保费();
			double 当月件数达成 = entry.getValue().get当月件数达成();
			double 当月保费达成 = entry.getValue().get当月保费达成();
			int 宽末未收件数 = entry.getValue().get宽末未收件数();
			int 宽一未收件数 = entry.getValue().get宽一未收件数();
			int 总未收件数 = entry.getValue().get总未收件数();

			if (课.equals("5课")) {
				总已收件数5课 = 总已收件数5课 + 当月已收件数;
				总应该收件数5课 = 总应该收件数5课 + 当月应收件数;

				try {
					sheet5课.addCell(new Label(0, count5, name));
					sheet5课.addCell(new Number(1, count5, 当月应收件数));
					sheet5课.addCell(new Number(2, count5, 当月应收保费));
					sheet5课.addCell(new Number(3, count5, 当月已收件数));
					sheet5课.addCell(new Number(4, count5, 当月已收保费));
					sheet5课.addCell(new Number(5, count5, 当月未收件数));
					sheet5课.addCell(new Number(6, count5, 当月未收保费));
					sheet5课.addCell(new Number(7, count5, 当月件数达成, wcfF));
					sheet5课.addCell(new Number(8, count5, 当月保费达成, wcfF));
					sheet5课.addCell(new Number(9, count5, 宽末未收件数));
					sheet5课.addCell(new Number(10, count5, 宽一未收件数));
					sheet5课.addCell(new Number(11, count5, 总未收件数));

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
					sheet11课.addCell(new Number(2, count11, 当月应收保费));
					sheet11课.addCell(new Number(3, count11, 当月已收件数));
					sheet11课.addCell(new Number(4, count11, 当月已收保费));
					sheet11课.addCell(new Number(5, count11, 当月未收件数));
					sheet11课.addCell(new Number(6, count11, 当月未收保费));
					sheet11课.addCell(new Number(7, count11, 当月件数达成, wcfF));
					sheet11课.addCell(new Number(8, count11, 当月保费达成, wcfF));
					sheet11课.addCell(new Number(9, count11, 宽末未收件数));
					sheet11课.addCell(new Number(10, count11, 宽一未收件数));
					sheet11课.addCell(new Number(11, count11, 总未收件数));

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
					sheet18课.addCell(new Number(2, count18, 当月应收保费));
					sheet18课.addCell(new Number(3, count18, 当月已收件数));
					sheet18课.addCell(new Number(4, count18, 当月已收保费));
					sheet18课.addCell(new Number(5, count18, 当月未收件数));
					sheet18课.addCell(new Number(6, count18, 当月未收保费));
					sheet18课.addCell(new Number(7, count18, 当月件数达成, wcfF));
					sheet18课.addCell(new Number(8, count18, 当月保费达成, wcfF));
					sheet18课.addCell(new Number(9, count18, 宽末未收件数));
					sheet18课.addCell(new Number(10, count18, 宽一未收件数));
					sheet18课.addCell(new Number(11, count18, 总未收件数));

				} catch (RowsExceededException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
				count18++;
			}
			else if (课.equals("22课")) {
				总已收件数22课 = 总已收件数22课 + 当月已收件数;
				总应该收件数22课 = 总应该收件数22课 + 当月应收件数;

				try {
					sheet22课.addCell(new Label(0, count22, name));
					sheet22课.addCell(new Number(1, count22, 当月应收件数));
					sheet22课.addCell(new Number(2, count22, 当月应收保费));
					sheet22课.addCell(new Number(3, count22, 当月已收件数));
					sheet22课.addCell(new Number(4, count22, 当月已收保费));
					sheet22课.addCell(new Number(5, count22, 当月未收件数));
					sheet22课.addCell(new Number(6, count22, 当月未收保费));
					sheet22课.addCell(new Number(7, count22, 当月件数达成, wcfF));
					sheet22课.addCell(new Number(8, count22, 当月保费达成, wcfF));
					sheet22课.addCell(new Number(9, count22, 宽末未收件数));
					sheet22课.addCell(new Number(10, count22, 宽一未收件数));
					sheet22课.addCell(new Number(11, count22, 总未收件数));

				} catch (RowsExceededException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
				count22++;
			}
		}

		整体达标5课 = 计算百分比(总已收件数5课, 总应该收件数5课, 2);
		整体达标11课 = 计算百分比(总已收件数11课, 总应该收件数11课, 2);
		整体达标18课 = 计算百分比(总已收件数18课, 总应该收件数18课, 2);
		整体达标22课 = 计算百分比(总已收件数22课, 总应该收件数22课, 2);
		try {
			sheet5课.addCell(new Label(0, 0, "5课当月件数整体达成" + 整体达标5课 + "%"));
			sheet11课.addCell(new Label(0, 0, "11课当月件数整体达成" + 整体达标11课 + "%"));
			sheet18课.addCell(new Label(0, 0, "18课当月件数整体达成" + 整体达标18课 + "%"));
			sheet22课.addCell(new Label(0, 0, "22课当月件数整体达成" + 整体达标22课 + "%"));

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

		System.out.println("生成excel :  <" + filepath+"> \n");
	}

	public String 计算百分比(double 已收, double 应收, int 保留位数) {
		String result;
		NumberFormat numberFormat = NumberFormat.getInstance();
		// 设置精确到小数点后2位
		numberFormat.setMaximumFractionDigits(2);
		result = numberFormat.format((float) 已收 / (float) 应收 * 100);
		return result;
	}

	public  LinkedHashMap<String, people> getAllPeople() {
		return allPeople;
	}

	public void setAllPeople( LinkedHashMap<String, people> allPeople) {
		this.allPeople = allPeople;
	}

}
