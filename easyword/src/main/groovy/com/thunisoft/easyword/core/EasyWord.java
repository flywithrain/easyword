package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * EasyWord
 *
 * @author 657518680@qq.com
 * @date 2019/8/13 9:34
 * @since 1.0.0
 */
public class EasyWord {

    /**
     * description
     *
     * @param inputStream  输入流
     * @param outputStream 输出流
     * @param staticLabel  静态标签
     * @param dynamicLabel 动态标签
     * @param tableLabel   列表标签
     * @param pictureLabel 图片标签
     * @author 657518680@qq.com
     * @date 2019/8/13
     * @since 1.0.0
     */
    public static void replaceLabel(@NotNull InputStream inputStream,
                                    @NotNull OutputStream outputStream,
                                    @NotNull Map<String, Customization> staticLabel,
                                    @NotNull Map<String, List<Customization>> dynamicLabel,
                                    @NotNull Map<String, List<List<Customization>>> tableLabel,
                                    @NotNull Map<String, Customization> pictureLabel)
            throws IOException, InvalidFormatException {
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        if (!staticLabel.isEmpty() || !dynamicLabel.isEmpty() || !tableLabel.isEmpty() || !pictureLabel.isEmpty()) {
            processParagraph(xwpfDocument, staticLabel, dynamicLabel, pictureLabel);
            processTable(xwpfDocument, staticLabel, tableLabel, pictureLabel);
        }
        xwpfDocument.write(outputStream);
    }

    /**
     * description
     *
     * @param xwpfDocument 文档
     * @param staticLabel  静态标签
     * @param dynamicLabel 动态标签
     * @param pictureLabel 图片标签
     * @author 657518680@qq.com
     * @date 2019/8/13
     * @since 1.0.0
     */
    private static void processParagraph(@NotNull XWPFDocument xwpfDocument,
                                         Map<String, Customization> staticLabel,
                                         Map<String, List<Customization>> dynamicLabel,
                                         Map<String, Customization> pictureLabel)
            throws IOException, InvalidFormatException {
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        pLable:
        for (int p = 0; p < paragraphs.size(); ++p) {
            XWPFParagraph paragraph = paragraphs.get(p);
            List<XWPFRun> runs = paragraph.getRuns();
            for (int r = 0; r < runs.size(); ++r) {
                XWPFRun run = runs.get(r);
                String text = run.text();
                Processor.processStaticLabel(staticLabel, paragraph, r, run, text);
                ProcessDynamicLabel4Paragraph processDynamicLabel4Paragraph
                        = new ProcessDynamicLabel4Paragraph(xwpfDocument, dynamicLabel, p, paragraph, run, text);
                processDynamicLabel4Paragraph.process();
                p = processDynamicLabel4Paragraph.getpIndex();
                if (processDynamicLabel4Paragraph.isContinue()) {
                    continue pLable;
                }
                Processor.processPicture4Paragraph(pictureLabel, paragraph, r, run, text);
            }
        }
    }

    /**
     * description
     *
     * @param xwpfDocument 文档
     * @param staticLabel  静态标签
     * @param tableLabel   列表标签
     * @param pictureLabel 图片标签
     * @author 657518680@qq.com
     * @date 2019/8/13
     * @since 1.0.0
     */
    private static void processTable(@NotNull XWPFDocument xwpfDocument,
                                     Map<String, Customization> staticLabel,
                                     Map<String, List<List<Customization>>> tableLabel,
                                     Map<String, Customization> pictureLabel)
            throws IOException, InvalidFormatException {
        List<XWPFTable> tables = xwpfDocument.getTables();
        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();
            rlabel:
            for (int r = 0; r < rows.size(); ++r) {
                XWPFTableRow row = rows.get(r);
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (int p = 0; p < paragraphs.size(); ++p) {
                        XWPFParagraph paragraph = paragraphs.get(p);
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (int i = 0; i < runs.size(); ++i) {
                            XWPFRun run = runs.get(i);
                            String text = run.text();
                            Processor.processStaticLabel(staticLabel, paragraph, i, run, text);
                            ProcessTable4Table processTable4Table =
                                    new ProcessTable4Table(tableLabel, table, r, row, paragraph, run, text);
                            processTable4Table.process();
                            r = processTable4Table.getRowIndex();
                            if (processTable4Table.isContinue()) {
                                continue rlabel;
                            }
                            Processor.processPicture4Table(pictureLabel, cell, text);
                        }
                    }
                }
            }
        }
    }


}
