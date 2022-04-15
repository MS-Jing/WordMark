package com.lj.wordmark.mark;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

/**
 * @Author luojing
 * @Date 2022/4/15
 */
@Slf4j
public abstract class AbstractWordMark {
    //文档对象
    protected final XWPFDocument document;
    //输入的word流
    private final InputStream inputStream;
    //输出流
    private OutputStream outputStream = null;

    public AbstractWordMark(InputStream in) throws IOException {
        document = new XWPFDocument(in);
        this.inputStream = in;
    }

    /**
     * 设置输出流
     *
     * @param filePath 输出文件路径
     * @throws FileNotFoundException 文本可能不存在的异常
     */
    public void setOutputFile(String filePath) throws FileNotFoundException {
        setOutputStream(new FileOutputStream(filePath));
    }

    /**
     * 设置输出流
     *
     * @param outputStream 输出流
     */
    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    /**
     * 用户设置了输出流后直接可以往输出流写出数据
     *
     * @throws IOException 写出异常
     */
    public void doWrite() throws IOException {
        //校验 输出流是否为空
        if (outputStream == null) {
            throw new RuntimeException("请先指定输出位置！");
        }
        try {
            //执行段落的模板替换
            log.info("执行段落的模板替换");
            doReplaceParagraph();
            //执行表格的模板替换
            log.info("执行表格的模板替换");
            doReplaceTable();
            //写出word
            log.info("开始写出word");
            writeWord();
            log.info("写出word结束");
        } finally {
            // 执行流的关闭
            close();
        }
    }

    /**
     * 执行段落的模板替换
     */
    protected abstract void doReplaceParagraph();

    /**
     * 执行表格的模板替换
     */
    protected abstract void doReplaceTable();

    /**
     * 真实写出word
     *
     * @throws IOException io流 异常
     */
    private void writeWord() throws IOException {
        document.write(outputStream);
    }

    /**
     * 关闭流操作
     * 其实document会自行帮我们关闭输入输出流
     *
     * @throws IOException io流 异常
     */
    private void close() throws IOException {
        inputStream.close();
        document.close();
        outputStream.close();
    }


}
