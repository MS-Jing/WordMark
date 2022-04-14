package com.lj.wordmark;

import com.lj.wordmark.utils.WordMarkUtils;

import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.StringJoiner;

/**
 * @Author luojing
 * @Date 2022/4/14
 */
public class Main {
    public static void main(String[] args) throws IOException {
        WordMark wordMark = new WordMark("test\\测试模板.docx");
        wordMark.outputFile("test\\测试wordMark生成.docx");
        Map<String, String> dataModel = new HashMap<>();
        dataModel.put("tableName","基础测试");
        dataModel.put("username","张三");
        dataModel.put("createTime",new Date().toString());
        wordMark.setDataModel(dataModel);
        wordMark.doWrite();
    }
}
