package com.lj.wordmark.mark;

import com.lj.wordmark.utils.MultiValueMap;
import com.lj.wordmark.utils.WordMarkUtils;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author luojing
 * @Date 2022/4/15
 */
@Slf4j
public class BaseWordMark extends AbstractWordMark {

    /**
     * 数据模型
     */
    @Setter
    private Map<String, String> dataModel = null;

    /**
     * 表格数据模型
     */
    @Setter
    private MultiValueMap<String, Map<String, String>> tableDataModel = null;
    /**
     * 表格标记行索引，默认第1行，从0开始
     */
    @Setter
    private Integer markRowIndex = 1;

    /**
     * 标记匹配
     */
    private static final Pattern MARK_COMPILE = Pattern.compile("\\$\\{(.*?)}");
    /**
     * 标记前缀
     */
    private static final String MARK_PREFIX = "${";
    /**
     * 标记后缀
     */
    private static final String MARK_SUFFIX = "}";

    private static final String TABLE_DELIMITER = ":";

    public BaseWordMark(String wordFilePath) throws IOException {
        this(new FileInputStream(wordFilePath));
    }

    public BaseWordMark(InputStream in) throws IOException {
        super(in);
    }

    @Override
    protected void doReplaceParagraph() {
        if (dataModel == null || dataModel.isEmpty()) {
            //如果没有数据模型则不用替换
            log.warn("没有设置相关的数据模型！dataModel");
            return;
        }
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        replaceParagraphMark(paragraphs, dataModel);
    }

    @Override
    protected void doReplaceTable() {
        if (tableDataModel == null || tableDataModel.isEmpty()) {
            //如果没有表格数据模型则不用替换
            log.warn("没有设置相关的数据模型！ tableDataModel");
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
                        //截取表格定界符来判断填充那张表
                        currentTableName = firstMark.split(TABLE_DELIMITER)[0];
                        tableDataList = tableDataModel.get(currentTableName);
                        break;
                    }
                }
            }

            //填充数据表
            if (tableDataList == null || tableDataList.isEmpty()) {
                return;
            }
            log.debug("填充表:" + currentTableName);
            for (Map<String, String> tableData : tableDataList) {
                //创建一个新的行
                XWPFTableRow dataRow = table.createRow();
                //获取这一行中的所有单元格
                List<XWPFTableCell> dataCells = dataRow.getTableCells();
                for (int i = 0; i < dataCells.size(); i++) {
                    XWPFTableCell markCell = markCells.get(i);
                    XWPFTableCell dataCell = dataCells.get(i);
                    WordMarkUtils.copyCell(markCell, dataCell);
                    replaceParagraphMark(dataCell.getParagraphs(), tableData, currentTableName + TABLE_DELIMITER);
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
            log.debug("替换段落:" + paragraph.getText());
            Set<XWPFRun> markRuns = extractMarks(paragraph);
            log.debug("提取出标记文本:" + markRuns);
            for (XWPFRun markRun : markRuns) {
                List<String> textMarks = resolverTextMark(markRun.text());
                if (textMarks.size() >= 1) {
                    String runContent = markRun.text();
                    runContent = runContent.replaceAll(modelPrefix, "");
                    for (String textMark : textMarks) {
                        textMark = textMark.replace(modelPrefix, "");
                        String data = model.getOrDefault(textMark, "");
                        //防止使用实体类的方式获取的数据本来就是null报空指针异常
                        runContent = runContent.replace(MARK_PREFIX + textMark + MARK_SUFFIX, data != null ? data : "");
                    }
                    log.debug("标记文本：{} 替换成：{}", markRun.text(), runContent);
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
    private Set<XWPFRun> extractMarks(XWPFParagraph paragraph) {
        Set<XWPFRun> markRuns = new HashSet<>();
        List<XWPFRun> runs = paragraph.getRuns();
        List<IndexRun> indexRuns = prepareIndexRuns(runs);

        String paragraphText = paragraph.getText();
        int startIndex = 0, endIndex = 0;
        while (true) {
            startIndex = WordMarkUtils.getSeparatorIndexOf(paragraphText, MARK_PREFIX, startIndex);
            if (startIndex != -1) {
                endIndex = WordMarkUtils.getSeparatorIndexOf(paragraphText, MARK_SUFFIX, startIndex + 1);
                if (endIndex != -1) {
                    //绝对不可能没有
                    List<IndexRun> indexRunRange = findIndexRunRange(indexRuns, startIndex, endIndex);
                    IndexRun lastIndexRun = indexRunRange.get(indexRunRange.size() - 1);
                    if (lastIndexRun.adjustAble) {
                        if (indexRunRange.size() > 1) {
                            lastIndexRun.forwardToAdjust();
                            //记得一定要把最后一个从indexRunRange删掉不然一会会从段落删除掉
                            indexRunRange.remove(indexRunRange.size() - 1);
                        } else {
                            lastIndexRun.backToAdjust();
                        }

                    }
                    Optional<IndexRun> reduce = indexRunRange.stream().reduce((f, l) -> {
                        f.getRun().setText(f.getRun().text() + l.getRun().text(), 0);
                        //将替换后的不完整的run置空，不然生成的文本会携带
                        l.getRun().setText("",0);
                        return f;
                    });
                    //将组合起来的标志位run添加到集合中
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
        return markRuns;
    }

    private List<IndexRun> prepareIndexRuns(List<XWPFRun> runs) {
        List<IndexRun> indexRuns = new ArrayList<>();
        int index = 0;
        IndexRun lastIndexRun = null;
        //组装索引run的列表
        for (XWPFRun run : runs) {
            int length = run.text().length();
            int lastIndex = index + length;
            IndexRun indexRun = new IndexRun(index, lastIndex, run, lastIndexRun);
            indexRuns.add(indexRun);
            index = lastIndex;
            lastIndexRun = indexRun;
        }
        //组装索引run的next
        for (int i = 0; i < indexRuns.size() - 1; i++) {
            indexRuns.get(i).setNext(indexRuns.get(i + 1));
        }

        return indexRuns;
    }

    /**
     * 从文本的索引列表找出在模板范围的文本索引
     *
     * @param indexRuns  索引列表
     * @param startIndex 开始索引
     * @param endIndex   结束索引
     * @return 在模板范围的文本索引
     */
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

    private static class IndexRun {
        private int startIndex;
        private int endIndex;
        @Getter
        private XWPFRun run;
        private IndexRun last;
        @Setter
        private IndexRun next;
        /**
         * 当run中}后面还有值表示可以调整
         */
        @Getter
        private boolean adjustAble = false;

        private IndexRun(int startIndex, int endIndex, XWPFRun run, IndexRun last) {
            this.startIndex = startIndex;
            this.endIndex = endIndex;
            this.run = run;
            this.last = last;
            //判断是否可以调整 }后面还有值表示不能删
            if (run.text().contains(MARK_SUFFIX) && !run.text().endsWith(MARK_SUFFIX)) {
                this.adjustAble = true;
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
         * 向前面调整，将}(包括})前面部分追加到上一个run
         */
        private void forwardToAdjust() {
            if (adjustAble && last != null) {
                String text = run.text();
                int i = text.indexOf(MARK_SUFFIX);
                String pre = text.substring(0, i + 1);
                String next = text.substring(i + 1);
                last.getRun().setText(last.getRun().text() + pre, 0);
                run.setText(next, 0);
                last.endIndex += pre.length();
                this.startIndex += pre.length();
            }
        }

        /**
         * 向后调整 将}(包括})后面的部分依附到下一个run
         */
        private void backToAdjust() {
            if (adjustAble && this.next != null) {
                String text = run.text();
                int i = text.indexOf(MARK_SUFFIX);
                String pre = text.substring(0, i + 1);
                String next = text.substring(i + 1);
                this.next.getRun().setText(next + this.next.getRun().text(), 0);
                run.setText(pre, 0);
                this.endIndex -= next.length();
                this.next.startIndex -= next.length();
            }
        }

        @Override
        public String toString() {
            return "IndexRun{" +
                    "startIndex=" + startIndex +
                    ", endIndex=" + endIndex +
                    ", run=" + run +
                    ", adjustAble=" + adjustAble +
                    '}';
        }
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
}
