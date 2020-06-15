import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class testExcel {

    public static void main(String args[]) throws EncryptedDocumentException, IOException, InvalidFormatException, ParseException {

        Map<String, Integer> weekly_complete = new HashMap<String, Integer>();
        Map<String, Integer> weekly_add = new HashMap<String, Integer>();
        Map<String, Integer> weekly_store = new HashMap<String, Integer>();
        String sunday = "2020-06-07";
        String saturday = "2020-06-13";
        DateFormat df2 = new SimpleDateFormat("yyyy-MM-dd");// HH:mm:ss
        Date sunday_date = df2.parse(sunday);
        Date saturday_date = df2.parse(saturday);
        String fileName = "C:\\Users\\xuesh\\Desktop\\XJD_综合业务系统-所有任务.xls";
        File xlsFile = new File(fileName);
        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);
        // 表个数
        int numberOfSheets = workbook.getNumberOfSheets();
//      System.out.println(numberOfSheets);
        //我们的需求只需要处理一个表，因此不需要遍历
        Sheet sheet = workbook.getSheetAt(0);
        // 行数
        int rowNumbers = sheet.getLastRowNum() + 1;

        for (int row = 1; row < rowNumbers - 1; row++) {
            Row r = sheet.getRow(row);
            String title = regexTitle(r.getCell(1).toString());
            String status = r.getCell(4).toString();
            Date create_time = r.getCell(6).getDateCellValue();
            Date create_date = df2.parse(df2.format(create_time));
            Date last_modify_time = df2.parse("2000-01-01");
            Date last_modify_date = df2.parse("2000-01-01");
            if (r.getCell(8).getCellType() == 0) {
                last_modify_time = r.getCell(8).getDateCellValue();
                last_modify_date = df2.parse(df2.format(last_modify_time));
            }
            if (status.equals("已完成") & last_modify_date.after(sunday_date)) {
                if (weekly_complete.containsKey(title)) {
                    Integer value = weekly_complete.get(title);
                    weekly_complete.put(title, value + 1);
                } else {
                    weekly_complete.put(title, 1);
                }
            }
            if (status.equals("未开始") & create_date.after(sunday_date)) {
                if (weekly_add.containsKey(title)) {
                    Integer value = weekly_add.get(title);
                    weekly_add.put(title, value + 1);
                } else {
                    weekly_add.put(title, 1);
                }
            }
            if (status.equals("未开始") | status.equals("进行中")) {
                if (weekly_store.containsKey(title)) {
                    Integer value = weekly_store.get(title);
                    weekly_store.put(title, value + 1);
                } else {
                    weekly_store.put(title, 1);
                }
            }
        }
        System.out.println("本周完成情况------------------");
        System.out.println(Arrays.toString(weekly_complete.entrySet().toArray()));
        writeToTxt("本周完成情况------------------",weekly_complete);
        System.out.println("本周新增------------------");
        System.out.println(Arrays.toString(weekly_add.entrySet().toArray()));
        writeToTxt("本周新增------------------",weekly_add);
        System.out.println("本周留存------------------");
        System.out.println(Arrays.toString(weekly_store.entrySet().toArray()));
        writeToTxt("本周留存------------------",weekly_store);
        System.out.println("写数据成功！");
    }

    public static String regexTitle(String title) {
        String[] titleList = title.split("/");
        if (titleList[1].length() > 0) {
            if (titleList[1].equals("其他")) {
                if (titleList[2].contains("(")) {
                    String[] title2 = titleList[2].split("\\(");
                    return title2[0];
                } else {
                    return titleList[2];
                }
            } else {
                if (titleList[1].contains("(")) {
                    String[] title2 = titleList[1].split("\\(");
                    return title2[0];
                } else {
                    return titleList[1];
                }
            }

        } else {
            String title_pr = "暂无数据";
            return title_pr;
        }

    }


    public static void writeToTxt(String title,Map<String, Integer> result){
        FileWriter fw = null;
        File file = new File("C:\\Users\\xuesh\\Desktop\\result.txt");
        try {
            if (!file.exists()) {
                file.createNewFile();
            }
            fw = new FileWriter(file,true);
            fw.write(title+"\r\n");
            String[] order = new String[]{"材料","产品","总体要求","完成目标","首页","数据反馈","事务","数据封存","收藏文档"};
            for(int i = 0;i< order.length;i++){
                if(result.containsKey(order[i])){
                    Integer value = result.get(order[i]);
                    String content = order[i]+" "+value;
                    fw.write(content+"\r\n");
                }
                else{
                    String content = order[i]+" "+0;
                    fw.write(content+"\r\n");
                }
            }
            fw.write("\r\n");
            fw.flush();

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }finally{
            if(fw != null){
                try {
                    fw.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
        }

    }
}
