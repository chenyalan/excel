package gzll;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by 陈亚兰 on 2018/3/28.
 */
public class RegTest {
    public static void main(String[] args){
      //  String reg="1989.09-";
        String reg="(19[0-9][0-9])\\.(0?[1-9]|1[0-9])-";
        Pattern p=Pattern.compile(reg);
        Matcher m=p.matcher("1973.01-2014.02我姐夫的开发  1984.03-");

//        m.groupCount();   //返回2,因为有2组
//        m.start(1);   //返回0 返回第一组匹配到的子字符串在字符串中的索引号
//        m.start(2);   //返回3
//        m.end(1);   //返回3 返回第一组匹配到的子字符串的最后一个字符在字符串中的索引位置.
//        m.end(2);   //返回7
//        m.group(1);   //返回aaa,返回第一组匹配到的子字符串
//        m.group(2);//返回2223,返回第二组匹配到的子字符串
//

        while(m.find()){
            System.out.print("start:"+m.start()+"  end:"+m.end()+"\n");
        }
    }
}
