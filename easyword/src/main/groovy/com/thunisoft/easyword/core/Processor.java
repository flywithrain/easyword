package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization4Picture;
import com.thunisoft.easyword.bo.Customization4Text;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import com.thunisoft.easyword.util.AnalyzeFileType;
import com.thunisoft.easyword.util.AnalyzeImageSize;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTPPrImpl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTRPrImpl;

import java.io.*;
import java.util.ArrayList;
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
 * @version 1.1.0
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
     * process the staticLabel for both paragraph and table in word
     *
     * @param staticLabel   staticLabel
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean processStaticLabel(Map<String, Customization4Text> staticLabel,
                                      WordConstruct wordConstruct,
                                      Index index) {
        XWPFRun run = wordConstruct.getRun();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        String text = run.text();
        for (Map.Entry<String, Customization4Text> entry : staticLabel.entrySet()) {
            int rIndex = index.getrIndex();
            Customization4Text customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text.trim())) {
                XWPFRun newRun = paragraph.insertNewRun(rIndex);
                CTRPr ctrPr = run.getCTR().getRPr();
                processVanish(ctrPr);
                newRun.getCTR().setRPr(ctrPr);
                paragraph.removeRun(rIndex + 1);
                newRun.setText(text.replace(key, customization.getText()));
                wordConstruct.setRun(newRun);
                customization.handle(wordConstruct, index);
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/8/19
     * process the dynamic label for paragraph
     *
     * @param xwpfDocument  xwpfDocument
     * @param dynamicLabel  dynamicLabel
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean processDynamicLabel4Paragraph(XWPFDocument xwpfDocument,
                                                 Map<String, List<Customization4Text>> dynamicLabel,
                                                 WordConstruct wordConstruct,
                                                 Index index) {
        XWPFRun run = wordConstruct.getRun();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        String text = run.text();
        int pIndex = index.getpIndex();
        for (Map.Entry<String, List<Customization4Text>> entry : dynamicLabel.entrySet()) {
            List<Customization4Text> customizationList = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text.trim())) {
                for (int i = 0; i < customizationList.size(); i++) {
                    Customization4Text customization = customizationList.get(i);
                    XmlCursor cursor = paragraph.getCTP().newCursor();
                    XWPFParagraph newPara = xwpfDocument.insertNewParagraph(cursor);
                    newPara.getCTP().setPPr(paragraph.getCTP().getPPr());
                    XWPFRun newRun = newPara.createRun();
                    newRun.getCTR().setRPr(run.getCTR().getRPr());
                    newRun.setText(customization.getText());
                    wordConstruct.setRun(newRun);
                    wordConstruct.setParagraph(newPara);
                    index.setrIndex(0);
                    index.setpIndex(pIndex + i);
                    customization.handle(wordConstruct, index);
                }
                xwpfDocument.removeBodyElement(xwpfDocument.getPosOfParagraph(paragraph));
                if(customizationList.isEmpty()){
                    index.setpIndex(pIndex - 1);
                }
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/8/19
     * process the table label (dynamic label in table) for table
     *
     * @param tableLabel    tableLabel
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean processTable4Table(Map<String, List<List<Customization4Text>>> tableLabel,
                                      WordConstruct wordConstruct,
                                      Index index) throws IOException, ClassNotFoundException {
        XWPFTable table = wordConstruct.getTable();
        XWPFTableRow row = wordConstruct.getRow();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rowIndex = index.getRowIndex();
        String text = run.text();
        for (Map.Entry<String, List<List<Customization4Text>>> entry : tableLabel.entrySet()) {
            String key = entry.getKey();
            List<List<Customization4Text>> listList = entry.getValue();
            if (key.equals(text)) {
                CTTrPr ctTrPr = row.getCtRow().getTrPr();
                CTPPr ctpPr = paragraph.getCTP().getPPr();
                CTPPr deepCopyPpr = deepClone((CTPPrImpl)ctpPr);
                String style = getTrPrString(ctTrPr);
                List<XWPFTableCell> tableCells = row.getTableCells();
                List<CTTcPr> ctTcPrList = new ArrayList<>();
                for (XWPFTableCell temp : tableCells) {
                    ctTcPrList.add(temp.getCTTc().getTcPr());
                }
                CTRPr deepCopyRpr = null;
                for (int j = 0; j < listList.size(); j++) {
                    List<Customization4Text> list = listList.get(j);
                    if(j == 0){
                        CTRPr ctrPr = run.getCTR().getRPr();
                        deepCopyRpr = deepClone((CTRPrImpl)ctrPr);
                    }
                    if (isTheNextRow(table, rowIndex, style, j)) {
                        XWPFTableRow newTableRow = table.getRow(rowIndex);
                        for (int k = 0; k < list.size(); k++) {
                            Customization4Text customization = list.get(k);
                            XWPFTableCell tableCell = newTableRow.getCell(k);
                            clearCell(tableCell);
                            XWPFParagraph xwpfParagraph = getFirstTableParagraph(tableCell);
                            xwpfParagraph.getCTP().setPPr(deepCopyPpr);
                            XWPFRun xwpfRun = xwpfParagraph.createRun();
                            xwpfRun.getCTR().setRPr(deepCopyRpr);
                            xwpfRun.setText(customization.getText());
                            wordConstruct.setParagraph(xwpfParagraph);
                            wordConstruct.setRun(xwpfRun);
                            index.setpIndex(0);
                            index.setrIndex(0);
                            customization.handle(wordConstruct, index);
                        }
                    } else {
                        XWPFTableRow newTableRow = table.insertNewTableRow(rowIndex);
                        for (int k = 0; k < tableCells.size(); k++) {
                            XWPFTableCell newTableCell = newTableRow.addNewTableCell();
                            if (k < list.size()) {
                                Customization4Text customization = list.get(k);
                                newTableCell.getCTTc().setTcPr(ctTcPrList.get(k));
                                XWPFParagraph newParagraph = getFirstTableParagraph(newTableCell);
                                newParagraph.getCTP().setPPr(deepCopyPpr);
                                XWPFRun newRun = newParagraph.createRun();
                                newRun.getCTR().setRPr(deepCopyRpr);
                                newRun.setText(customization.getText());
                                wordConstruct.setRow(newTableRow);
                                wordConstruct.setCell(newTableCell);
                                wordConstruct.setParagraph(newParagraph);
                                wordConstruct.setRun(newRun);
                                index.setcIndex(k);
                                index.setpIndex(0);
                                index.setrIndex(0);
                                customization.handle(wordConstruct, index);
                            }
                        }
                        newTableRow.getCtRow().setTrPr(ctTrPr);
                    }
                    ++rowIndex;
                }
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/9/30 18:10
     * process the vertical label (dynamic label in table) for table
     *
     * @param verticalLabel vertical label
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author wangxiaoyu 657518680@qq.com
     * @since 1.1.0
     */
    static boolean processVerticalLabel(Map<String, List<Customization4Text>> verticalLabel,
                                        WordConstruct wordConstruct,
                                        Index index) throws IOException, ClassNotFoundException {
        XWPFTable table = wordConstruct.getTable();
        XWPFTableRow row = wordConstruct.getRow();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rowIndex = index.getRowIndex();
        int originalRowIndex = rowIndex;
        int cellIndex = index.getcIndex();
        String text = run.text();
        for (Map.Entry<String, List<Customization4Text>> entry : verticalLabel.entrySet()) {
            String key = entry.getKey();
            List<Customization4Text> list = entry.getValue();
            if (key.equals(text)) {
                CTTrPr ctTrPr = row.getCtRow().getTrPr();
                CTPPr ctpPr = paragraph.getCTP().getPPr();
                CTPPr deepCopyPpr = deepClone((CTPPrImpl)ctpPr);
                CTRPr ctrPr = run.getCTR().getRPr();
                CTRPr deepCopyRpr = deepClone((CTRPrImpl)ctrPr);
                processVanish(deepCopyRpr);
                String style = getTrPrString(ctTrPr);
                List<XWPFTableCell> tableCells = row.getTableCells();
                List<CTTcPr> ctTcPrList = new ArrayList<>();
                for (XWPFTableCell temp : tableCells) {
                    ctTcPrList.add(temp.getCTTc().getTcPr());
                }
                for (int i = 0; i < list.size(); i++) {
                    Customization4Text customization = list.get(i);
                    if (isTheNextRow(table, rowIndex, style, i)) {
                        XWPFTableRow tempRow = table.getRow(rowIndex);
                        XWPFTableCell tempCell = tempRow.getCell(cellIndex);
                        clearCell(tempCell);
                        XWPFParagraph tempParagraph = getFirstTableParagraph(tempCell);
                        tempParagraph.getCTP().setPPr(deepCopyPpr);
                        XWPFRun tempRun = tempParagraph.createRun();
                        tempRun.getCTR().setRPr(deepCopyRpr);
                        tempRun.setText(customization.getText());
                        wordConstruct.setParagraph(tempParagraph);
                        wordConstruct.setRun(tempRun);
                        index.setRowIndex(rowIndex);
                        index.setpIndex(0);
                        index.setrIndex(0);
                        customization.handle(wordConstruct, index);
                    } else {
                        XWPFTableRow tempRow = table.insertNewTableRow(rowIndex);
                        for (int k = 0; k < tableCells.size(); k++) {
                            XWPFTableCell tempCell = tempRow.addNewTableCell();
                            tempCell.getCTTc().setTcPr(ctTcPrList.get(k));
                            if (k == cellIndex) {
                                XWPFParagraph tempParagraph = getFirstTableParagraph(tempCell);
                                tempParagraph.getCTP().setPPr(deepCopyPpr);
                                XWPFRun tempRun = tempParagraph.createRun();
                                tempRun.getCTR().setRPr(deepCopyRpr);
                                tempRun.setText(customization.getText());
                                wordConstruct.setRow(tempRow);
                                wordConstruct.setCell(tempCell);
                                wordConstruct.setParagraph(tempParagraph);
                                wordConstruct.setRun(tempRun);
                                index.setRowIndex(rowIndex);
                                index.setpIndex(0);
                                index.setrIndex(0);
                                customization.handle(wordConstruct, index);
                            }
                        }
                    }
                    rowIndex++;
                }
                index.setRowIndex(--originalRowIndex);
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/8/19
     * process the picture label for paragraph
     *
     * @param pictureLabel  pictureLabel
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since alpha
     */
    static boolean processPicture4All(Map<String, Customization4Picture> pictureLabel,
                                      WordConstruct wordConstruct,
                                      Index index) throws IOException, InvalidFormatException {
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rIndex = index.getrIndex();
        String text = run.text();
        for (Map.Entry<String, Customization4Picture> entry : pictureLabel.entrySet()) {
            Customization4Picture customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text)) {
                XWPFRun newRun = paragraph.insertNewRun(rIndex);
                CTRPr ctrPr = run.getCTR().getRPr();
                processVanish(ctrPr);
                newRun.getCTR().setRPr(ctrPr);
                paragraph.removeRun(rIndex + 1);
                wordConstruct.setRun(newRun);
                processPicture(customization, newRun);
                customization.handle(wordConstruct, index);
                return true;
            }
        }
        return false;
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
     * 2019/8/20 14:15
     *
     * @param customization the all info of picture
     * @param newRun        the run created to save image
     * @author 657518680@qq.com
     * @since beta
     */
    private static void processPicture(Customization4Picture customization,
                                       XWPFRun newRun) throws IOException, InvalidFormatException {
        byte[] bytes = IOUtils.toByteArray(customization.getPicture());
        int width = customization.getWidth();
        int height = customization.getHeight();
        if (width <= 0 || height <= 0) {
            Map<String, Integer> size = AnalyzeImageSize.getImageSize(new ByteArrayInputStream(bytes));
            width = size.get("width");
            height = size.get("height");
        }
        newRun.addPicture(new ByteArrayInputStream(bytes),
                AnalyzeFileType.getFileType(bytes),
                customization.getPictureName(),
                Units.pixelToEMU(width),
                Units.pixelToEMU(height));
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
     * get the vanish attribute and set the value to STOnOff.FALSE {@link STOnOff#FALSE}
     *
     * @param ctrPr the attribute of the run of the word
     * @author 657518680@qq.com
     * @since alpha
     */
    private static void processVanish(CTRPr ctrPr) {
        if (ctrPr != null) {
            CTOnOff vanish = ctrPr.getVanish();
            if (vanish != null && !vanish.isSetVal()) {
                vanish.setVal(STOnOff.FALSE);
            }
        }
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
    private static XWPFParagraph getFirstTableParagraph(XWPFTableCell tableCell) {
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
    private static void clearCell(XWPFTableCell tableCell) {
        for (int i = tableCell.getParagraphs().size() - 1; i >= 0; i--) {
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
    private static boolean isTheNextRow(XWPFTable table, int rowIndex, String style, int j) {
        return j == 0 || (table.getRow(rowIndex) != null
                && style.equals(getTrPrString(table.getRow(rowIndex).getCtRow().getTrPr())));
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
    private static String getTrPrString(CTTrPr ctTrPr) {
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
    private static <T extends Serializable> T deepClone(T obj) throws IOException, ClassNotFoundException {
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
