package com.lj.wordmark.utils;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.List;

/**
 * @Author luojing
 * @Date 2022/4/13
 */
public class WordMarkUtils {

    /**
     * 复制文本run，包括复制样式
     *
     * @param source 来源run
     * @param target 目标run
     */
    public static void copyRun(XWPFRun source, XWPFRun target) {
        copyRun(source, target, source == null ? null : source.text());
    }

    /**
     * 复制文本run，包括复制样式
     *
     * @param source 来源run
     * @param target 目标run
     * @param text   目标run的文本
     */
    public static void copyRun(XWPFRun source, XWPFRun target, String text) {
        if (source == null || target == null) {
            throw new RuntimeException("复制run时source和target都不得为空！");
        }
        //设置文本
        target.setText(text, 0);
        //设置加粗
        if (!target.isBold()) {
            target.setBold(source.isBold());
        }
        //设置斜体
        target.setItalic(source.isItalic());
        //下划线
        target.setUnderline(source.getUnderline());
        //字体颜色
        target.setColor(source.getColor());
        //文本位置
        target.setTextPosition(source.getTextPosition());
        //字体大小
        if (source.getFontSize() != -1) {
            target.setFontSize(source.getFontSize());
        }
        //字体样式
        if (source.getFontFamily() != null) {
            target.setFontFamily(source.getFontFamily());
        }
        // 底下我也不知道设置了些啥
//        if (source.getCTR() != null && source.getCTR().isSetRPr() && source.getCTR().getRPr().isSetRFonts()) {
//            CTFonts sourceFonts = source.getCTR().getRPr().getRFonts();
//            CTRPr targetRPr = target.getCTR().isSetRPr() ? target.getCTR().getRPr() : target.getCTR().addNewRPr();
//            CTFonts targetFonts = targetRPr.isSetRFonts() ? targetRPr.getRFonts() : targetRPr.addNewRFonts();
//
//            //设置编码
//            targetFonts.setAscii(sourceFonts.getAscii());
//            targetFonts.setAsciiTheme(sourceFonts.getAsciiTheme());
//            targetFonts.setCs(sourceFonts.getCs());
//            targetFonts.setCstheme(sourceFonts.getCstheme());
//            targetFonts.setEastAsia(sourceFonts.getEastAsia());
//            targetFonts.setEastAsiaTheme(sourceFonts.getEastAsiaTheme());
//            targetFonts.setHAnsi(sourceFonts.getHAnsi());
//            targetFonts.setHAnsiTheme(sourceFonts.getHAnsiTheme());
//        }
    }

    /**
     * 复制段落包括样式和 文本内容
     *
     * @param source 来段落
     * @param target 模板段落
     */
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
        if (source == null || target == null) {
            throw new RuntimeException("复制Paragraph时source和target都不得为空！");
        }
        //复制段落的所有Run
        List<XWPFRun> sourceRuns = source.getRuns();
        if (!sourceRuns.isEmpty()) {
            sourceRuns.forEach(sr -> {
                XWPFRun targetRun = target.createRun();
                copyRun(sr, targetRun);
            });
        }
        // 复制段落信息
        target.setAlignment(source.getAlignment());
        target.setVerticalAlignment(source.getVerticalAlignment());
        target.setBorderBetween(source.getBorderBetween());
        target.setBorderBottom(source.getBorderBottom());
        target.setBorderLeft(source.getBorderLeft());
        target.setBorderRight(source.getBorderRight());
        target.setBorderTop(source.getBorderTop());
        target.setPageBreak(source.isPageBreak());

        if (source.getCTP() != null && source.getCTP().getPPr() != null) {
            CTPPr sourcePPr = source.getCTP().getPPr();
            CTPPr targetPPr = target.getCTP().getPPr() != null ? target.getCTP().getPPr() : target.getCTP().addNewPPr();
            //复制段落间距信息
            CTSpacing sourceSpacing = sourcePPr.getSpacing();
            if (sourceSpacing != null) {
                CTSpacing targetSpacing = targetPPr.getSpacing() != null ? targetPPr.getSpacing() : targetPPr.addNewSpacing();
                if (sourceSpacing.getAfter() != null) {
                    targetSpacing.setAfter(sourceSpacing.getAfter());
                }
                if (sourceSpacing.getAfterAutospacing() != null) {
                    targetSpacing.setAfterAutospacing(sourceSpacing
                            .getAfterAutospacing());
                }
                if (sourceSpacing.getAfterLines() != null) {
                    targetSpacing.setAfterLines(sourceSpacing.getAfterLines());
                }
                if (sourceSpacing.getBefore() != null) {
                    targetSpacing.setBefore(sourceSpacing.getBefore());
                }
                if (sourceSpacing.getBeforeAutospacing() != null) {
                    targetSpacing.setBeforeAutospacing(sourceSpacing
                            .getBeforeAutospacing());
                }
                if (sourceSpacing.getBeforeLines() != null) {
                    targetSpacing.setBeforeLines(sourceSpacing.getBeforeLines());
                }
                if (sourceSpacing.getLine() != null) {
                    targetSpacing.setLine(sourceSpacing.getLine());
                }
                if (sourceSpacing.getLineRule() != null) {
                    targetSpacing.setLineRule(sourceSpacing.getLineRule());
                }
            }
            // 复制段落缩进信息
            CTInd sourceInd = sourcePPr.getInd();
            if (sourceInd != null) {
                CTInd targetInd = targetPPr.getInd() != null ? targetPPr.getInd() : targetPPr.addNewInd();
                if (sourceInd.getFirstLine() != null) {
                    targetInd.setFirstLine(sourceInd.getFirstLine());
                }
                if (sourceInd.getFirstLineChars() != null) {
                    targetInd.setFirstLineChars(sourceInd.getFirstLineChars());
                }
                if (sourceInd.getHanging() != null) {
                    targetInd.setHanging(sourceInd.getHanging());
                }
                if (sourceInd.getHangingChars() != null) {
                    targetInd.setHangingChars(sourceInd.getHangingChars());
                }
                if (sourceInd.getLeft() != null) {
                    targetInd.setLeft(sourceInd.getLeft());
                }
                if (sourceInd.getLeftChars() != null) {
                    targetInd.setLeftChars(sourceInd.getLeftChars());
                }
                if (sourceInd.getRight() != null) {
                    targetInd.setRight(sourceInd.getRight());
                }
                if (sourceInd.getRightChars() != null) {
                    targetInd.setRightChars(sourceInd.getRightChars());
                }
            }
        }
    }

    /**
     * 复制单元格
     *
     * @param source 来源cell
     * @param target 模板cell
     */
    public static void copyCell(XWPFTableCell source, XWPFTableCell target) {
        if (source == null || target == null) {
            throw new RuntimeException("复制Cell时source和target都不得为空！");
        }

        CTTc sourceCttc = source.getCTTc();
        CTTcPr sourceTcPr = sourceCttc.getTcPr();

        CTTc targetCttc = target.getCTTc();
        CTTcPr targetTcPr = targetCttc.addNewTcPr();

        //复制单元格宽度
        if (sourceTcPr.getTcW() != null) {
            targetTcPr.addNewTcW().setW(sourceTcPr.getTcW().getW());
        }
        //复制单元格对齐方式
        if (sourceTcPr.getVAlign() != null) {
            targetTcPr.addNewVAlign().setVal(sourceTcPr.getVAlign().getVal());
        }
        if (sourceCttc.getPList().size() > 0) {
            CTP ctp = sourceCttc.getPList().get(0);
            if (ctp.getPPr() != null && ctp.getPPr().getJc() != null) {
                targetCttc.getPList().get(0).addNewPPr().addNewJc().setVal(ctp.getPPr().getJc().getVal());
            }
        }
        //复制边框
        if (sourceTcPr.getTcBorders() != null) {
            targetTcPr.setTcBorders(sourceTcPr.getTcBorders());
        }
        //复制段落
        List<XWPFParagraph> sourceParagraphs = source.getParagraphs();
        if (!sourceParagraphs.isEmpty()) {
            List<XWPFParagraph> targetParagraphs = target.getParagraphs();
            int targetsize = targetParagraphs.size();
            for (int i = 0; i < sourceParagraphs.size(); i++) {
                XWPFParagraph targetParagraph;
                if (i < targetsize) {
                    targetParagraph = targetParagraphs.get(i);
                } else {
                    targetParagraph = target.addParagraph();
                }
                copyParagraph(sourceParagraphs.get(i), targetParagraph);
            }
        }
    }

    /**
     * 从fromIndex开始获取分隔符的位置，
     * 如果不存在返回-1
     * 如果存在返回索引
     */
    public static int getSeparatorIndexOf(String name, String delimiter, int fromIndex) {
        return name.indexOf(delimiter, fromIndex);
    }
}
