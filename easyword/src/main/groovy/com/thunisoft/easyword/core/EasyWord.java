package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

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
 * @since alpha
 * @version beta
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
     * @since alpha
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
     * @since alpha
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
     * 2019/8/19 21:53
     * merge words
     *
     * @param wordList     the list of inputStream of word that need to be merge to one
     * @param outputStream the word that merged by the wordList
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     * @throws XmlException           XmlException
     * @author 657518680@qq.com
     * @since beta
     */
    public static void mergeWord(List<InputStream> wordList, OutputStream outputStream)
            throws IOException, InvalidFormatException, XmlException {
        if (CollectionUtils.isEmpty(wordList)) {
            return;
        }
        XWPFDocument newDocument = null;
        CTBody newCtBody = null;
        String newString = null;
        String prefix = null;
        StringBuilder mainPart = new StringBuilder();
        for (int i = 0; i < wordList.size(); ++i) {
            try (InputStream word = wordList.get(i)) {
                XWPFDocument xwpfDocument = new XWPFDocument(word);
                if (i != wordList.size() - 1) {
                    XWPFRun run = xwpfDocument.getLastParagraph().createRun();
                    run.addBreak(BreakType.PAGE);
                }
                CTBody ctBody = xwpfDocument.getDocument().getBody();
                if (i == 0) {
                    newDocument = xwpfDocument;
                    newCtBody = ctBody;
                    newString = newCtBody.xmlText();
                    prefix = newString.substring(0, newString.indexOf('>') + 1);
                    mainPart.append(newString, newString.indexOf('>') + 1, newString.lastIndexOf('<'));
                } else {
                    Processor.mergeOther2First(newDocument, mainPart, xwpfDocument, ctBody);
                }
            }
        }
        String sufix = null;
        if (newString != null) {
            sufix = newString.substring(newString.lastIndexOf('<'));
        }
        if (newCtBody != null) {
            newCtBody.set(CTBody.Factory.parse(prefix + mainPart.toString() + sufix));
        }
        if (newDocument != null) {
            newDocument.write(outputStream);
        }
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
     * @since alpha
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
     * @since alpha
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
