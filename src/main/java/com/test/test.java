package com.test;

import com.poi.ExcelExportUtil;

import java.util.ArrayList;
import java.util.List;

/**
 * @author YangQing
 * @version 1.0.0
 */

public class test {
    public static void main(String[] args) {
        Person Person1=new Person(1l,"1");
        Person Person2=new Person(2l,"2");
        Person Person3=new Person(3l,"3");
        Person Person4=new Person(4l,"4");
        Person Person5=new Person(5l,"5");
        Person Person6=new Person(6l,"6");
        List<Person> ps=new ArrayList<Person>();
        ps.add(Person1);
        ps.add(Person2);
        ps.add(Person3);
        ps.add(Person4);
        ps.add(Person5);
        ps.add(Person6);
        String[] headers = {"用户编号", "用户名称"};
        //表字段
        String[] fields = {"id", "name"};

        String msg=new ExcelExportUtil().exportUtil("11111",headers,fields,ps);
        System.out.println(msg);
    }
}
