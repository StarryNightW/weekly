import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class hello_test {
    public  static void main(String args[]){
        /*String test_title = "/其他(#3032)";
        String[] titleList = test_title.split("\\(");
        System.out.println(titleList[0]);
        System.out.println(titleList[1]);*/

        Map<String, Integer> weekly_complete = new HashMap<String,Integer>();

        weekly_complete.put("已完成",1);
        Integer value = weekly_complete.get("已完成");

        weekly_complete.put("已完成",value+1);
        weekly_complete.put("未完成",2);
        System.out.println(Arrays.toString(weekly_complete.entrySet().toArray()));
        String[] order = new String[]{"材料","产品","总体要求","完成目标","首页","数据反馈","事务","数据封存","收藏文档列表"};
        for(int i = 0;i< order.length;i++){
            System.out.println(order[i]);
        }



    }
}
