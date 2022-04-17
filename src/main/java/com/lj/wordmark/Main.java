package com.lj.wordmark;

import com.lj.wordmark.entity.DataModel;
import com.lj.wordmark.entity.Phone;
import com.lj.wordmark.entity.User;
import com.lj.wordmark.mark.BaseWordMark;
import com.lj.wordmark.mark.GeneralWordMark;
import com.lj.wordmark.utils.LinkedMultiValueMap;

import java.io.IOException;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;


/**
 * @Author luojing
 * @Date 2022/4/13
 */
public class Main {
    public static void main(String[] args) throws IOException {
//        //基础的测试
//        BaseWordMark wordMark = new BaseWordMark("test\\测试模板.docx");
//        wordMark.setOutputFile("test\\测试WordMark生成.docx");
//        // 创建数据模型
//        Map<String, String> map = new HashMap<>();
//        map.put("tableName", "测试用户");
//        map.put("username", "张三");
//        map.put("createTime", "2022/4/15");
//        wordMark.setDataModel(map);
//        //创建表格数据模型
//        LinkedMultiValueMap<String, Map<String, String>> tableDataModel = new LinkedMultiValueMap<>();
//        Map<String, String> userTable1 = new HashMap<>();
//        userTable1.put("index", "1");
//        userTable1.put("name", "张三");
//        userTable1.put("age", "18");
//        tableDataModel.add("user", userTable1);
//        Map<String, String> userTable2 = new HashMap<>();
//        userTable2.put("index", "2");
//        userTable2.put("name", "李四");
//        userTable2.put("age", "19");
//        userTable2.put("hobby", "篮球");
//        tableDataModel.add("user", userTable2);
//        Map<String, String> userTable3 = new HashMap<>();
//        userTable3.put("index", "3");
//        userTable3.put("name", "王五");
//        userTable3.put("age", "20");
//        tableDataModel.add("user", userTable3);
//        wordMark.setTableDataModel(tableDataModel);
//        //导出word
//        wordMark.doWrite();


        //==================================
        GeneralWordMark<DataModel, User> wordMark = new GeneralWordMark<>("test\\测试模板.docx");
        wordMark.setOutputFile("test\\测试WordMark生成.docx");
        // 创建数据模型
        DataModel dataModel = new DataModel();
        dataModel.setTableName("测试实体类");
        dataModel.setUsername("李四");
        dataModel.setCreateTime(LocalDate.now());
        wordMark.setDataModel(dataModel);
        //创建表格数据模型
        LinkedMultiValueMap<String, User> tableDataModel = new LinkedMultiValueMap<>();
        tableDataModel.add("user", new User(1, "张三", 18, null, null));
        tableDataModel.add("user", new User(2, "李四", 19, null, null));
        tableDataModel.add("user", new User(3, "王五", 20, "羽毛球", new Phone("123", "小米")));
        wordMark.setObjectTableDataModel(tableDataModel);

        //导出word
        wordMark.doWrite();

    }
}
