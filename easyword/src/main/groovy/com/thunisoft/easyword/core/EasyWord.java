package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.*;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.jetbrains.annotations.NotNull;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
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
     * 2019/8/24 14:58
     * a simplified version of {@link EasyWord#replaceLabel(InputStream, OutputStream, Map)}
     *
     * @param inputStream    inputStream
     * @param outputStream   outputStream
     * @param staticLabelite a simplified version of staticLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static void replaceLabelite(@NotNull InputStream inputStream,
                                       @NotNull OutputStream outputStream,
                                       @NotNull Map<String, String> staticLabelite)
            throws IOException, InvalidFormatException, ClassNotFoundException {
        replaceLabel(inputStream, outputStream, staticLite2Full(staticLabelite));
    }

    /**
     * 2019/8/24 14:48
     * a simplified version of {@link EasyWord#replaceLabel(InputStream, OutputStream, Map, Map, Map, Map, Map)}
     *
     * @param inputStream     inputStream
     * @param outputStream    outputStream
     * @param staticLabelite  a simplified version of staticLabel
     * @param dynamicLabelite a simplified version of dynamicLabel
     * @param tableLabelite   a simplified version of tableLabel
     * @param pictureLabel    pictureLabel
     * @param verticalLabel   verticalLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static void replaceLabelite(@NotNull InputStream inputStream,
                                       @NotNull OutputStream outputStream,
                                       @NotNull Map<String, String> staticLabelite,
                                       @NotNull Map<String, List<String>> dynamicLabelite,
                                       @NotNull Map<String, List<List<String>>> tableLabelite,
                                       @NotNull Map<String, Customization4Picture> pictureLabel,
                                       @NotNull Map<String, List<String>> verticalLabel)
            throws IOException, InvalidFormatException, ClassNotFoundException {
        replaceLabel(inputStream,
                outputStream,
                staticLite2Full(staticLabelite),
                dynamicLite2Full(dynamicLabelite),
                tableLite2Full(tableLabelite),
                pictureLabel,
                dynamicLite2Full(verticalLabel));
    }

    /**
     * 2019/8/19
     * replace the label in the word
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
                                    @NotNull Map<String, Customization4Text> staticLabel)
            throws IOException, InvalidFormatException, ClassNotFoundException {
        replaceLabel(inputStream, outputStream, staticLabel,
                new HashMap<>(0),
                new HashMap<>(0),
                new HashMap<>(0),
                new HashMap<>(0));
    }

    /**
     * 2019/8/13
     * replace the label in the word
     *
     * @param inputStream   inputStream
     * @param outputStream  outputStream
     * @param staticLabel   staticLabel
     * @param dynamicLabel  dynamicLabel
     * @param tableLabel    tableLabel
     * @param pictureLabel  pictureLabel
     * @param verticalLabel verticalLabel
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     * @author 657518680@qq.com
     * @since alpha
     */
    public static void replaceLabel(@NotNull InputStream inputStream,
                                    @NotNull OutputStream outputStream,
                                    @NotNull Map<String, Customization4Text> staticLabel,
                                    @NotNull Map<String, List<Customization4Text>> dynamicLabel,
                                    @NotNull Map<String, List<List<Customization4Text>>> tableLabel,
                                    @NotNull Map<String, Customization4Picture> pictureLabel,
                                    @NotNull Map<String, List<Customization4Text>> verticalLabel)
            throws IOException, InvalidFormatException, ClassNotFoundException {
        XWPFDocument xwpfDocument = new XWPFDocument(inputStream);
        if (!staticLabel.isEmpty() || !dynamicLabel.isEmpty() || !tableLabel.isEmpty()
                || !pictureLabel.isEmpty() || !verticalLabel.isEmpty()) {
            processParagraph(xwpfDocument, staticLabel, dynamicLabel, pictureLabel);
            processTable(xwpfDocument, staticLabel, tableLabel, pictureLabel, verticalLabel);
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
        for (Map.Entry xmlns : headMap.entrySet()) {
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
     * @param staticLabel  staticLabel
     * @param dynamicLabel dynamicLabel
     * @param pictureLabel pictureLabel
     * @author 657518680@qq.com
     * @since alpha
     */
    private static void processParagraph(@NotNull XWPFDocument xwpfDocument,
                                         Map<String, Customization4Text> staticLabel,
                                         Map<String, List<Customization4Text>> dynamicLabel,
                                         Map<String, Customization4Picture> pictureLabel)
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
                    flag = Processor.processPicture4All(pictureLabel, wordConstruct, index);
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
                                     Map<String, Customization4Text> staticLabel,
                                     Map<String, List<List<Customization4Text>>> tableLabel,
                                     Map<String, Customization4Picture> pictureLabel,
                                     Map<String, List<Customization4Text>> verticalLabel)
            throws IOException, InvalidFormatException, ClassNotFoundException {
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
                                flag = Processor.processPicture4All(pictureLabel, wordConstruct, index);
                            }
                            boolean result = false;
                            if (!flag) {
                                result = Processor.processTable4Table(tableLabel, wordConstruct, index);
                            }
                            if (!result) {
                                result = Processor.processVerticalLabel(verticalLabel, wordConstruct, index);
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

    /**
     * 2019/8/24 14:48
     * Convert staticLabelite to staticLabel
     *
     * @param staticLabelite a simplified version of staticLabel
     * @return staticLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static Map<String, Customization4Text> staticLite2Full(Map<String, String> staticLabelite) {
        Map<String, Customization4Text> staticLabel = new HashMap<>(staticLabelite.size());
        for (Map.Entry<String, String> entry : staticLabelite.entrySet()) {
            staticLabel.put(entry.getKey(), new DefaultCustomization(entry.getValue()));
        }
        return staticLabel;
    }

    /**
     * 2019/8/24 14:48
     * Convert dynamicLabelite to dynamicLabel
     *
     * @param dynamicLabelite a simplified version of dynamicLabel
     * @return dynamicLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static Map<String, List<Customization4Text>> dynamicLite2Full(Map<String, List<String>> dynamicLabelite) {
        Map<String, List<Customization4Text>> dynamicLabel = new HashMap<>(dynamicLabelite.size());
        for (Map.Entry<String, List<String>> entry : dynamicLabelite.entrySet()) {
            List<Customization4Text> temp = new ArrayList<>(entry.getValue().size());
            entry.getValue().forEach(str -> temp.add(new DefaultCustomization(str)));
            dynamicLabel.put(entry.getKey(), temp);
        }
        return dynamicLabel;
    }

    /**
     * 2019/8/24 14:48
     * Convert tableLabelite to tableLabel
     *
     * @param tableLabelite a simplified version of tableLabel
     * @return tableLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static Map<String, List<List<Customization4Text>>>
    tableLite2Full(Map<String, List<List<String>>> tableLabelite) {
        Map<String, List<List<Customization4Text>>> tableLabel = new HashMap<>(tableLabelite.size());
        for (Map.Entry<String, List<List<String>>> entry : tableLabelite.entrySet()) {
            List<List<Customization4Text>> rows = new ArrayList<>(entry.getValue().size());
            entry.getValue().forEach((List<String> list) -> {
                List<Customization4Text> row = new ArrayList<>(list.size());
                list.forEach(str -> row.add(new DefaultCustomization(str)));
                rows.add(row);
            });
            tableLabel.put(entry.getKey(), rows);
        }
        return tableLabel;
    }

}
