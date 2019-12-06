package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 2019/8/13 19:07
 * Processor of EasyWord
 *
 * @author 657518680@qq.com
 * @version 2.0.0
 * @since alpha
 */
final class Processor {

    /**
     * 2019/8/23 11:03
     * <p>
     * get xml xmlns
     *
     * @since beta
     */
    private static final Pattern PATTERN = Pattern.compile("(xmlns(:[\\s\\S]+?)?)=[\\s\\S]+?\\s");

    private Processor() {
    }

    /**
     * 2019/8/19
     * process the label for both paragraph and table in word
     *
     * @param label   label
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean processLabel(Map<String, Customization> label,
                                WordConstruct wordConstruct,
                                Index index) {
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        for (Map.Entry<String, Customization> entry : label.entrySet()) {
            String key = entry.getKey();
            TextSegment textSegment = paragraph.searchText(key, new PositionInParagraph());
            if (textSegment != null) {
                handleRunInParagraph(wordConstruct, index, textSegment);
                Customization customization = entry.getValue();
                customization.handle(key, wordConstruct, index);
                return true;
            }
        }
        return false;
    }

    private static void handleRunInParagraph(WordConstruct wordConstruct, Index index, TextSegment textSegment) {
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        List<XWPFRun> runs = paragraph.getRuns();
        int beginRun = textSegment.getBeginRun();
        int endRun = textSegment.getEndRun();
        StringBuilder b = new StringBuilder();
        XWPFRun tempRun;
        for (int runPos = beginRun; runPos <= endRun; runPos++) {
            tempRun = runs.get(runPos);
            b.append(tempRun.text());
            clearRun(tempRun);
        }
        XWPFRun run = runs.get(beginRun);
        run.setText(b.toString());
        wordConstruct.setRun(run);
        index.setrIndex(beginRun);
    }

    static XWPFRun clearRun(XWPFRun run){
        XWPFParagraph paragraph = (XWPFParagraph) run.getParent();
        int runIndex = paragraph.getRuns().indexOf(run);
        CTRPr ctrPr = run.getCTR().getRPr();
        XWPFRun newRun = paragraph.insertNewRun(runIndex);
        newRun.getCTR().setRPr(ctrPr);
        paragraph.removeRun(runIndex + 1);
        return newRun;
    }

    /**
     * 2019/9/30 18:11
     *
     * @param head    document.xml file header
     * @param headMap the map of head
     * @author wangxiaoyu 657518680@qq.com
     * @since beta
     */
    static void getXmlns(String head, Map<String, Object> headMap) {
        Matcher matcher = PATTERN.matcher(head);
        while (matcher.find()) {
            headMap.put(matcher.group(1), matcher.group());
        }
    }

    /**
     * 2019/8/19
     * get the vanish attribute and set the value to STOnOff.FALSE {@link STOnOff#FALSE}
     *
     * @param ctrPr the attribute of the run of the word
     * @author 657518680@qq.com
     * @since alpha
     */
    static void processVanish(CTRPr ctrPr) {
        if (ctrPr != null) {
            CTOnOff vanish = ctrPr.getVanish();
            if (vanish != null && !vanish.isSetVal()) {
                vanish.setVal(STOnOff.FALSE);
            }
        }
    }

    /**
     * 2019/8/19 21:53
     * Merge the word other than the first word into the first word
     *
     * @param newDocument  the document of the first word and will create the final word
     * @param mainPart     the main part of the first word and will keep adding mainPart of other words
     * @param xwpfDocument the document of the word that except the first
     * @param ctBody       the body of the word that except the first
     * @throws InvalidFormatException InvalidFormatException
     * @author 657518680@qq.com
     * @since beta
     */
    static void mergeOther2First(XWPFDocument newDocument,
                                 StringBuilder mainPart,
                                 XWPFDocument xwpfDocument,
                                 CTBody ctBody,
                                 Map<String, Object> headMap) throws InvalidFormatException {
        XmlOptions xmlOptions = new XmlOptions();
        xmlOptions.setSaveOuter();
        String appendString = ctBody.xmlText(xmlOptions);
        getXmlns(appendString.substring(1, appendString.indexOf('>')) + " ", headMap);
        String addPart = appendString
                .substring(appendString.indexOf('>') + 1, appendString.lastIndexOf('<'));
        List<XWPFPictureData> allPictures = xwpfDocument.getAllPictures();
        if (allPictures != null) {
            // 记录图片合并前及合并后的ID
            Map<String, String> map = new HashMap<>();
            for (XWPFPictureData picture : allPictures) {
                String before = xwpfDocument.getRelationId(picture);
                //将原文档中的图片加入到目标文档中
                String after = newDocument.addPictureData(picture.getData(), picture.getPictureType());
                map.put(before, after);
            }
            if (!map.isEmpty()) {
                //对xml字符串中图片ID进行替换
                for (Map.Entry<String, String> set : map.entrySet()) {
                    addPart = addPart.replace(set.getKey(), set.getValue());
                }
            }
        }
        mainPart.append(addPart);
    }

    /**
     * 2019/8/19
     * get the first paragraph of the cell if not exist then create a new paragraph
     *
     * @param tableCell the cell of table
     * @return the first paragraph
     * @author 657518680@qq.com
     * @since alpha
     */
    static XWPFParagraph getFirstTableParagraph(XWPFTableCell tableCell) {
        List<XWPFParagraph> paragraphList = tableCell.getParagraphs();
        if (paragraphList.isEmpty()) {
            return tableCell.addParagraph();
        }
        return paragraphList.get(0);
    }

    /**
     * 2019/9/30 18:13
     * clear all in the cell
     *
     * @param tableCell the table cell
     * @author wangxiaoyu 657518680@qq.com
     * @since 1.0.0
     */
    static void clearCell(XWPFTableCell tableCell) {
        for (int i = tableCell.getParagraphs().size() - 1; ; i--) {
            if(i < 0){
                return;
            }
            tableCell.removeParagraph(i);
        }
    }

    /**
     * 2019/8/19
     * determine if the row is the next row that can be used for table label rather than create a new row
     *
     * @param table    table
     * @param rowIndex rowIndex {@link Index#getRowIndex()}
     * @param style    the string of the style{@linkplain Processor#getTrPrString}
     * @param j        the index of the list of the value of the table label
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean isTheNextRow(XWPFTable table, int rowIndex, String style, int j) {
        return j == 0 || (table.getRow(rowIndex) != null
                && style.equals(getTrPrString(table.getRow(rowIndex).getCtRow().getTrPr()))
                && table.getRow(rowIndex).getTableCells().size()
                == table.getRow(rowIndex - 1).getTableCells().size());
    }

    /**
     * 2019/8/19
     * get the string of ctTrPr if null then return empty string
     *
     * @param ctTrPr the attribute of the run of the word
     * @return the string of the ctTrPr
     * @author 657518680@qq.com
     * @since alpha
     */
    static String getTrPrString(CTTrPr ctTrPr) {
        if (ctTrPr == null) {
            return "";
        }
        return ctTrPr.toString();
    }

    /**
     * 2019/9/30 18:55
     *
     * @param obj the object need to deep copy
     * @return the deep copy
     * @author wangxiaoyu 657518680@qq.com
     * @since 1.1.0
     */
    static <T extends Serializable> T deepClone(T obj) throws IOException, ClassNotFoundException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ObjectOutputStream obs = new ObjectOutputStream(out);
        obs.writeObject(obj);
        obs.close();
        ByteArrayInputStream ios = new ByteArrayInputStream(out.toByteArray());
        ObjectInputStream ois = new ObjectInputStream(ios);
        T cloneObj = (T) ois.readObject();
        ois.close();
        return cloneObj;
    }

}
