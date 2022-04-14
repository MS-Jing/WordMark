package com.lj.wordmark;

import com.lj.wordmark.utils.MultiValueMap;
import com.lj.wordmark.utils.WordMarkUtils;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author luojing
 * @Date 2022/4/13
 * word的模板替换工具
 * 替换段落文本，替换表格内容
 */
@Slf4j
public class WordMark {
    //文档对象
    private final XWPFDocument document;
    //输入的word流
    private final InputStream inputStream;
    //输出流
    private OutputStream outputStream = null;
    //数据模型
    @Setter
    private Map<String, String> dataModel = null;
    //表格数据模型
    @Setter
    private MultiValueMap<String, Map<String, String>> tableDataModel = null;
    // 标记匹配
    private static final Pattern MARK_COMPILE = Pattern.compile("\\$\\{(.*?)}");
    // 表格标记行索引，默认第1行，从0开始
    @Setter
    private Integer markRowIndex = 1;

    public WordMark(String wordFilePath) throws IOException {
        this(new FileInputStream(wordFilePath));
    }

    public WordMark(InputStream in) throws IOException {
        document = new XWPFDocument(in);
        this.inputStream = in;
    }

    public void outputFile(String filePath) throws FileNotFoundException {
        setOutputStream(new FileOutputStream(filePath));
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public void doWrite() throws IOException {
        //校验 输出流是否为空
        if (outputStream == null) {
            throw new RuntimeException("请先指定输出位置！");
        }
        try {
            //执行段落的模板替换
            doReplaceParagraph();
            //表格的生成
            doReplaceTable();
            //写出word
            writeWord();
        } finally {
            // 执行流的关闭
            close();
        }
    }

    /**
     * 段落文本替换
     */
    private void doReplaceParagraph() {
        if (dataModel == null || dataModel.isEmpty()) {
            //如果没有数据模型则不用替换
            return;
        }
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        replaceParagraphMark(paragraphs, dataModel);
    }

    /**
     * 替换和生成表格
     */
    private void doReplaceTable() {
        if (tableDataModel == null || tableDataModel.isEmpty()) {
            //如果没有表格数据模型则不用替换
            return;
        }
        // 获取文档中的所有表格
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            //获取当前表格的所有行
            List<XWPFTableRow> rows = table.getRows();
            //标记的行，就是除了表头的第一行(这是必须的！！！)
            List<XWPFTableCell> markCells = null;
            // 填充数据的数据表列表
            List<Map<String, String>> tableDataList = null;
            // 当前数据表名
            String currentTableName = null;

            // 遍历约定好的标记行（这里是第一行）的单元格，确定填充哪一个数据表
            if (rows.size() >= markRowIndex + 1) {
                //标记的行
                XWPFTableRow markRow = rows.get(markRowIndex);
                //标记的单元格
                markCells = markRow.getTableCells();
                //遍历每一个标记的单元格，如果遇到第一个标记来确定当前表使用那个表数据模型填充
                for (XWPFTableCell markCell : markCells) {
                    List<String> markList = resolverTextMark(markCell.getText());
                    if (!markList.isEmpty()) {
                        String firstMark = markList.get(0);
                        currentTableName = firstMark.split("\\.")[0];
                        tableDataList = tableDataModel.get(currentTableName);
                        break;
                    }
                }
            }

            //填充数据表
            if (tableDataList == null || tableDataList.isEmpty()) {
                return;
            }
            if (markCells == null) {
                throw new RuntimeException("标记行必须在第二行！");
            }
            for (Map<String, String> tableData : tableDataList) {
                //创建一个新的行
                XWPFTableRow dataRow = table.createRow();
                //获取这一行中的所有单元格
                List<XWPFTableCell> dataCells = dataRow.getTableCells();
                for (int i = 0; i < dataCells.size(); i++) {
                    XWPFTableCell markCell = markCells.get(i);
                    XWPFTableCell dataCell = dataCells.get(i);
                    WordMarkUtils.copyCell(markCell, dataCell);
                    replaceParagraphMark(dataCell.getParagraphs(), tableData, currentTableName + ".");
                }
            }

            //填充结束删除标记行
            table.removeRow(markRowIndex);
        }

    }

    /**
     * 替换段落中的标记占位符
     * 因为一个单元格也可以看成多个段落，所以这块代码可以进行复用
     *
     * @param paragraphs 段落列表
     * @param model      数据模型映射
     */
    private void replaceParagraphMark(List<XWPFParagraph> paragraphs, Map<String, String> model) {
        replaceParagraphMark(paragraphs, model, "");
    }

    /**
     * 替换段落中的标记占位符
     *
     * @param paragraphs  段落列表
     * @param model       数据模型映射
     * @param modelPrefix 数据模型的前缀，像表格会有前缀`user.name` model已经是对应表了所以需要去掉对应的前缀
     */
    private void replaceParagraphMark(List<XWPFParagraph> paragraphs, Map<String, String> model, String modelPrefix) {
        if (model == null || model.isEmpty()) {
            //如果没有数据模型则不用替换
            return;
        }

        //循环每个段落,获取所有标志位，解析标志位替换文本
        for (XWPFParagraph paragraph : paragraphs) {
            Set<XWPFRun> markRuns = extractMarks(paragraph);
            for (XWPFRun markRun : markRuns) {
                List<String> textMarks = resolverTextMark(markRun.text());
                if (textMarks.size() >= 1) {
                    String runContent = markRun.text();
                    runContent = runContent.replaceAll(modelPrefix, "");
                    for (String textMark : textMarks) {
                        textMark = textMark.replace(modelPrefix, "");
                        String data = model.getOrDefault(textMark, "");
                        runContent = runContent.replace("${" + textMark + "}", data);
                    }
                    markRun.setText(runContent, 0);
                } else {
                    log.warn(("请确认你的标志位的语法是否正确:" + markRun.text()));
                }
            }
        }
    }

    /**
     * 提取段落中的标志位
     *
     * @param paragraph 段落
     * @return 提取后的标志位，一个段落可以有多个标志位
     */
//    private List<XWPFRun> extractMarks(XWPFParagraph paragraph) {
//        List<XWPFRun> markRuns = new ArrayList<>();
//        //段落的所有文本
//        List<XWPFRun> runs = paragraph.getRuns();
//        //当匹配到开始标志位时存放标志位，直到处理到结束标志位
//        Deque<XWPFRun> deque = new ArrayDeque<>();
//        Iterator<XWPFRun> iterator = runs.iterator();
//        //会从开始标志位拷贝到结束标志位，所以除了结束标志位都是不完整的标志位
//        HashSet<XWPFRun> removeSet = new HashSet<>();
//        while (iterator.hasNext()) {
//            XWPFRun run = iterator.next();
//            if (run.text().startsWith("$")) {
//                deque.add(run);
//                while (!run.text().endsWith("}")) {
//                    run = iterator.next();
//                    deque.add(run);
//                }
//                //开始组装完整的标记位
//                if (deque.peekLast().text().endsWith("}")) {
//                    Optional<XWPFRun> reduce = deque.stream().reduce((f, l) -> {
//                        WordMarkUtils.copyRun(f, l, f.text() + l.text());
//                        removeSet.add(f);
//                        return l;
//                    });
//                    if (reduce.isPresent()) {
//                        XWPFRun markRun = reduce.get();
//                        markRuns.add(markRun);
//                    }
//                    // 在一个段落中一个标志位替换完毕一定要将队列清空
//                    deque.clear();
//                } else {
//                    throw new RuntimeException("匹配替换发生错误！在 " + run.text() + " 附近。你必须以 `}` 结束但是以 " + deque.peekLast().text() + " 结束");
//                }
//            }
//        }
//        //删除当前段落不完整的run
//        removeSet.forEach(e -> {
//            int i = runs.indexOf(e);
//            paragraph.removeRun(i);
//        });
//        return markRuns;
//    }

    /**
     * //$在中间：判断${
     * //是：判断是否有}没有就一直找}结束
     * //否：清空队列下一个
     * //$在结尾,需要判断下一个是否是{开头
     * //是：判断是否有}没有就一直找}结束
     * //否: 清空队列下一个
     *
     * @param paragraph
     * @return
     */
    private Set<XWPFRun> extractMarks(XWPFParagraph paragraph) {
        Set<XWPFRun> markRuns = new HashSet<>();
        List<XWPFRun> runs = paragraph.getRuns();
        Set<XWPFRun> removeSet = new HashSet<>();
        List<IndexRun> indexRuns = new ArrayList<>();
        int index = 0;
        for (XWPFRun run : runs) {
            int length = run.text().length();
            int lastIndex = index + length;
            indexRuns.add(new IndexRun(index, lastIndex, run));
            index = lastIndex;
        }
        String paragraphText = paragraph.getText();
        int startIndex = 0, endIndex = 0;
        while (true) {
            startIndex = WordMarkUtils.getSeparatorIndexOf(paragraphText, "${", startIndex);
            if (startIndex != -1) {
                endIndex = WordMarkUtils.getSeparatorIndexOf(paragraphText, "}", startIndex + 1);
                if (endIndex != -1) {
                    //绝对不可能没有
                    List<IndexRun> indexRunRange = findIndexRunRange(indexRuns, startIndex, endIndex);
                    IndexRun lastIndexRun = indexRunRange.get(indexRunRange.size() - 1);
                    if (!lastIndexRun.removeAble) {
                        //倒数第二个
                        IndexRun penultIndexRun = indexRunRange.get(indexRunRange.size() - 2);
                        //最后一个会自己调整，将}前面包括}的标志位内容返回
                        // 由倒数第二个run进行添加
                        String adjust = lastIndexRun.adjust();
                        penultIndexRun.getRun().setText(adjust);
                    }
                    Optional<IndexRun> reduce = indexRunRange.stream().reduce((f, l) -> {
                        f.getRun().setText(l.getRun().text());
//                        WordMarkUtils.copyRun(l.getRun(), f.getRun(), f.getRun().text() + l.getRun().text());
                        removeSet.add(l.getRun());
                        return f;
                    });
                    reduce.ifPresent(indexRun -> markRuns.add(indexRun.getRun()));
                    startIndex = endIndex + 1;
                } else {
                    //有开始符没有结束符是语法错误
                    throw new RuntimeException("出现语法错误！在索引 " + startIndex + " 出没有与之匹配的结束符");
                }
            } else {
                //说明该段结束了或者没有找到${开始符
                break;
            }
        }
        //删除当前段落不完整的run
        removeSet.forEach(e -> {
            int i = runs.indexOf(e);
            paragraph.removeRun(i);
        });
        return markRuns;
    }

    private List<IndexRun> findIndexRunRange(List<IndexRun> indexRuns, int startIndex, int endIndex) {
        List<IndexRun> indexRunRanges = new ArrayList<>();
        Iterator<IndexRun> iterator = indexRuns.iterator();
        while (iterator.hasNext()) {
            IndexRun next = iterator.next();
            if (next.thereof(startIndex)) {
                while (!next.thereof(endIndex)) {
                    indexRunRanges.add(next);
                    next = iterator.next();
                }
                //记得把最后一个加上
                indexRunRanges.add(next);
            }
        }
        return indexRunRanges;
    }

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

    /**
     * 解析文本的内容的标记
     *
     * @param text 内容
     * @return 标记的字符 例如: [user.name,user.age]
     */
    private List<String> resolverTextMark(String text) {
        List<String> markList = new ArrayList<>();
        if (text == null || "".equals(text)) {
            return markList; //如果该列文本是空的返回一个空的List
        }
        Matcher matcher = MARK_COMPILE.matcher(text);
        //循环匹配防止文本中有多个标记占位
        while (matcher.find()) {
            markList.add(matcher.group(1));
        }
        return markList;
    }

    @ToString
    private static class IndexRun {
        private int startIndex;
        private int endIndex;
        @Getter
        private XWPFRun run;
        /**
         * 可以删除不代表一定要删
         * 当run中}后面还有值表示不能删
         */
        @Getter
        private boolean removeAble = true;

        private IndexRun(int startIndex, int endIndex, XWPFRun run) {
            this.startIndex = startIndex;
            this.endIndex = endIndex;
            this.run = run;
            //判断是否可以删除 }后面还有值表示不能删
            if (run.text().contains("}") && !run.text().endsWith("}")) {
                this.removeAble = false;
            }
        }

        /**
         * 判断字符索引是否在当前run中
         * 左包含右不包含
         *
         * @param index 字符索引
         * @return 在其中返回true
         */
        private boolean thereof(int index) {
            return startIndex <= index && index < endIndex;
        }

        /**
         * 调整，将当前run }后面的保留，前面的返回
         *
         * @return }前面的返回包括}
         */
        private String adjust() {
            if (!removeAble) {
                String text = run.text();
                int i = text.indexOf("}");
                String pre = text.substring(0, i + 1);
                String next = text.substring(i + 1);
                WordMarkUtils.copyRun(run, run, next);
                return pre;
            }
            return "";
        }
    }
}
