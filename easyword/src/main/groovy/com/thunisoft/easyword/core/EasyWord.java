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
 * @version beta
 * @since alpha
 */
public final class EasyWord {

    private static final String HEAD = "<xml-fragment ";

    private EasyWord() {

    }

    /**
     * 2019/8/13
     * replace the label in the word
     *
     * @param inputStream  inputStream
     * @param outputStream outputStream
     * @param label        label
     * @throws IOException IOException
     * @author 657518680@qq.com
     * @since alpha
     */
    public static void replaceLabel(@NotNull InputStream inputStream,
                                    @NotNull OutputStream outputStream,
                                    @NotNull Map<String, Customization> label) throws IOException {
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        if (!label.isEmpty()) {
            processParagraph(xwpfDocument, label);
            processTable(xwpfDocument, label);
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
        String newString;
        Map<String, Object> headMap = new HashMap<>(30);
        String sufix = null;
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
                    Processor.getXmlns(newString.substring(0, newString.indexOf('>')) + " ", headMap);
                    mainPart.append(newString, newString.indexOf('>') + 1, newString.lastIndexOf('<'));
                    sufix = newString.substring(newString.lastIndexOf('<'));
                } else {
                    Processor.mergeOther2First(newDocument, mainPart, xwpfDocument, ctBody, headMap);
                }
            }
        }
        StringBuilder prefix = new StringBuilder(HEAD);
        for (Map.Entry<String, Object> xmlns : headMap.entrySet()) {
            prefix.append(xmlns.getValue());
        }
        prefix.append(">");
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
     * @param label        staticLabel
     * @author 657518680@qq.com
     * @since alpha
     */
    private static void processParagraph(@NotNull XWPFDocument xwpfDocument, Map<String, Customization> label) {
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        for (int p = 0; p < paragraphs.size(); ++p) {
            XWPFParagraph paragraph = paragraphs.get(p);
            Index index = new Index(p, 0);
            WordConstruct wordConstruct = new WordConstruct(xwpfDocument, paragraph, null);
            if (Processor.processLabel(label, wordConstruct, index)) {
                p = index.getpIndex();
            }
        }
    }

    /**
     * 2019/8/13
     * description
     *
     * @param xwpfDocument xwpfDocument
     * @param label        label
     * @author 657518680@qq.com
     * @since alpha
     */
    private static void processTable(@NotNull XWPFDocument xwpfDocument, Map<String, Customization> label) {
        List<XWPFTable> tables = xwpfDocument.getTables();
        for (int t = 0; t < tables.size(); ++t) {
            XWPFTable table = tables.get(t);
            List<XWPFTableRow> rows = table.getRows();
            for (int r = 0; r < rows.size(); ++r) {
                XWPFTableRow row = rows.get(r);
                List<XWPFTableCell> cells = row.getTableCells();
                for (int c = 0; c < cells.size(); ++c) {
                    XWPFTableCell cell = cells.get(c);
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (int p = 0; p < paragraphs.size(); ++p) {
                        XWPFParagraph paragraph = paragraphs.get(p);
                        Index index = new Index(t, r, c, p, 0);
                        WordConstruct wordConstruct =
                                new WordConstruct(xwpfDocument, table, row, cell, paragraph, null);
                        if (Processor.processLabel(label, wordConstruct, index)) {
                            t = index.getTableIndex();
                            r = index.getRowIndex();
                            c = index.getcIndex();
                            p = index.getpIndex();
                        }
                    }
                }
            }
        }
    }

}
