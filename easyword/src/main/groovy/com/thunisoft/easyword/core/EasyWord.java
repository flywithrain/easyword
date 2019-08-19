package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 2019/8/13 9:34
 * EasyWord
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public final class EasyWord {

    private EasyWord() {

    }

    /**
     * 2019/8/19
     * description
     *
     * @param inputStream  inputStream
     * @param outputStream outputStream
     * @param staticLabel  staticLabel
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static void replaceLabel(@NotNull InputStream inputStream,
                                    @NotNull OutputStream outputStream,
                                    @NotNull Map<String, Customization> staticLabel)
            throws IOException, InvalidFormatException {
        replaceLabel(inputStream, outputStream, staticLabel,
                new HashMap<>(0),
                new HashMap<>(0),
                new HashMap<>(0));
    }

    /**
     * 2019/8/13
     * description
     *
     * @param inputStream  inputStream
     * @param outputStream outputStream
     * @param staticLabel  staticLabel
     * @param dynamicLabel dynamicLabel
     * @param tableLabel   tableLabel
     * @param pictureLabel pictureLabel
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     * @author 657518680@qq.com
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
     * 2019/8/13
     * description
     *
     * @param xwpfDocument xwpfDocument
     * @param staticLabel  staticLabel
     * @param dynamicLabel dynamicLabel
     * @param pictureLabel pictureLabel
     * @author 657518680@qq.com
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
                Index index = new Index(p, r);
                WordConstruct wordConstruct = new WordConstruct(paragraph, run);
                //是否已经处理过run
                boolean flag = Processor.processStaticLabel(staticLabel, wordConstruct, index);
                if (!flag) {
                    flag = Processor.processPicture4Paragraph(pictureLabel, wordConstruct, index);
                }
                boolean result = false;
                if (!flag) {
                    result = Processor.processDynamicLabel4Paragraph(xwpfDocument, dynamicLabel, wordConstruct, index);
                }
                p = index.getpIndex();
                r = index.getrIndex();
                if (result) {
                    continue pLable;
                }
            }
        }
    }

    /**
     * 2019/8/13
     * description
     *
     * @param xwpfDocument xwpfDocument
     * @param staticLabel  staticLabel
     * @param tableLabel   tableLabel
     * @param pictureLabel pictureLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    private static void processTable(@NotNull XWPFDocument xwpfDocument,
                                     Map<String, Customization> staticLabel,
                                     Map<String, List<List<Customization>>> tableLabel,
                                     Map<String, Customization> pictureLabel)
            throws IOException, InvalidFormatException {
        List<XWPFTable> tables = xwpfDocument.getTables();
        for (int t = 0; t < tables.size(); ++t) {
            XWPFTable table = tables.get(t);
            List<XWPFTableRow> rows = table.getRows();
            rlabel:
            for (int r = 0; r < rows.size(); ++r) {
                XWPFTableRow row = rows.get(r);
                List<XWPFTableCell> cells = row.getTableCells();
                for (int c = 0; c < cells.size(); ++c) {
                    XWPFTableCell cell = cells.get(c);
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (int p = 0; p < paragraphs.size(); ++p) {
                        XWPFParagraph paragraph = paragraphs.get(p);
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (int i = 0; i < runs.size(); ++i) {
                            XWPFRun run = runs.get(i);
                            Index index = new Index(t, r, c, p, i);
                            WordConstruct wordConstruct = new WordConstruct(table, row, cell, paragraph, run);
                            boolean flag = Processor.processStaticLabel(staticLabel, wordConstruct, index);
                            if (!flag) {
                                flag = Processor.processPicture4Table(pictureLabel, wordConstruct, index);
                            }
                            boolean result = false;
                            if (!flag) {
                                result = Processor.processTable4Table(tableLabel, wordConstruct, index);
                            }
                            t = index.getTableIndex();
                            r = index.getRowIndex();
                            c = index.getcIndex();
                            p = index.getpIndex();
                            i = index.getrIndex();
                            if (result) {
                                continue rlabel;
                            }
                        }
                    }
                }
            }
        }
    }

}
