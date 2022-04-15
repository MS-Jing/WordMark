package com.lj.wordmark;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * @Author luojing
 * @Date 2022/4/15
 */
public class TestMain {
    public static void main(String[] args) throws IOException {
        XWPFDocument document = new XWPFDocument(new FileInputStream("test\\测试WordMark生成.docx"));
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            for (XWPFRun run : paragraph.getRuns()) {
                System.out.println(run);
            }
        }
        document.close();
    }
}
