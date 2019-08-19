package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import com.thunisoft.easyword.util.AnalyzeFileType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 2019/8/13 19:07
 * Processor of EasyWord
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
final class Processor {

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
     * @since 1.0.0
     */
    static boolean processStaticLabel(Map<String, Customization> staticLabel,
                                      WordConstruct wordConstruct,
                                      Index index) {
        XWPFRun run = wordConstruct.getRun();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        String text = run.text();
        for (Map.Entry<String, Customization> entry : staticLabel.entrySet()) {
            int rIndex = index.getrIndex();
            Customization customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text.trim())) {
                XWPFRun newRun = paragraph.insertNewRun(rIndex);
                CTRPr ctrPr = run.getCTR().getRPr();
                processVanish(ctrPr);
                newRun.getCTR().setRPr(ctrPr);
                paragraph.removeRun(rIndex + 1);
                newRun.setText(text.replaceAll(key, customization.getText()));
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
     * @since 1.0.0
     */
    static boolean processDynamicLabel4Paragraph(XWPFDocument xwpfDocument,
                                                 Map<String, List<Customization>> dynamicLabel,
                                                 WordConstruct wordConstruct,
                                                 Index index) {
        XWPFRun run = wordConstruct.getRun();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        String text = run.text();
        for (Map.Entry<String, List<Customization>> entry : dynamicLabel.entrySet()) {
            List<Customization> customizationList = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text.trim())) {
                for (Customization customization : customizationList) {
                    XmlCursor cursor = paragraph.getCTP().newCursor();
                    XWPFParagraph newPara = xwpfDocument.insertNewParagraph(cursor);
                    newPara.getCTP().setPPr(paragraph.getCTP().getPPr());
                    XWPFRun newRun = newPara.createRun();
                    newRun.getCTR().setRPr(run.getCTR().getRPr());
                    newRun.setText(customization.getText());
                    customization.handle(wordConstruct, index);
                }
                xwpfDocument.removeBodyElement(xwpfDocument.getPosOfParagraph(paragraph));
                int pIndex = index.getpIndex();
                pIndex += customizationList.size();
                index.setpIndex(pIndex);
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
     * @since 1.0.0
     */
    static boolean processTable4Table(Map<String, List<List<Customization>>> tableLabel,
                                      WordConstruct wordConstruct,
                                      Index index) {
        XWPFTable table = wordConstruct.getTable();
        XWPFTableRow row = wordConstruct.getRow();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rowIndex = index.getRowIndex();
        String text = run.text();
        for (Map.Entry<String, List<List<Customization>>> entry : tableLabel.entrySet()) {
            String key = entry.getKey();
            List<List<Customization>> listList = entry.getValue();
            if (key.equals(text)) {
                CTTrPr ctTrPr = row.getCtRow().getTrPr();
                String style = getTrPrString(ctTrPr);
                List<XWPFTableCell> tableCells = row.getTableCells();
                List<CTTcPr> ctTcPrList = new ArrayList<>();
                for (XWPFTableCell temp : tableCells) {
                    ctTcPrList.add(temp.getCTTc().getTcPr());
                }
                int temp = rowIndex;
                for (int j = 0; j < listList.size(); j++) {
                    List<Customization> list = listList.get(j);
                    CTRPr ctrPr = run.getCTR().getRPr();
                    Processor.processVanish(ctrPr);
                    XWPFTableRow newTableRow;
                    if (isHasNextRow(table, rowIndex, style, j)) {
                        newTableRow = table.getRow(rowIndex + 1);
                        for (int k = 0; k < list.size(); k++) {
                            Customization customization = list.get(k);
                            XWPFTableCell tableCell = newTableRow.getCell(k);
                            XWPFParagraph xwpfParagraph = getFirstTableParagraph(tableCell);
                            XWPFRun xwpfRun = xwpfParagraph.createRun();
                            xwpfRun.getCTR().setRPr(ctrPr);
                            xwpfRun.setText(customization.getText());
                            customization.handle(wordConstruct, index);
                        }
                        ++rowIndex;
                    } else {
                        newTableRow = table.insertNewTableRow(rowIndex + 1);
                        for (int k = 0; k < list.size(); k++) {
                            Customization customization = list.get(k);
                            XWPFTableCell newTableCell = newTableRow.addNewTableCell();
                            newTableCell.getCTTc().setTcPr(ctTcPrList.get(k));
                            XWPFParagraph newParagraph = getFirstTableParagraph(newTableCell);
                            newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                            XWPFRun newRun = newParagraph.createRun();
                            newRun.getCTR().setRPr(ctrPr);
                            newRun.setText(customization.getText());
                            customization.handle(wordConstruct, index);
                        }
                        newTableRow.getCtRow().setTrPr(ctTrPr);
                        ++rowIndex;
                    }
                }
                --rowIndex;
                table.removeRow(temp);
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
     * @since 1.0.0
     */
    static boolean processPicture4Paragraph(Map<String, Customization> pictureLabel,
                                            WordConstruct wordConstruct,
                                            Index index) throws IOException, InvalidFormatException {
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rIndex = index.getrIndex();
        String text = run.text();
        for (Map.Entry<String, Customization> entry : pictureLabel.entrySet()) {
            Customization customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text)) {
                XWPFRun newRun = paragraph.insertNewRun(rIndex);
                CTRPr ctrPr = run.getCTR().getRPr();
                processVanish(ctrPr);
                newRun.getCTR().setRPr(ctrPr);
                paragraph.removeRun(rIndex + 1);
                newRun.addPicture(customization.getPicture(),
                        AnalyzeFileType.getFileType(customization.getPicture()),
                        customization.getPictureName(),
                        Units.toEMU(customization.getWidth()),
                        Units.toEMU(customization.getHeight()));
                customization.handle(wordConstruct, index);
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/8/19
     * process the picture label for table
     *
     * @param pictureLabel  pictureLabel
     * @param wordConstruct wordConstruct {@link WordConstruct}
     * @param index         index{@link Index}
     * @return true: already processed; false: not processed
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    static boolean processPicture4Table(Map<String, Customization> pictureLabel,
                                        WordConstruct wordConstruct,
                                        Index index) throws IOException, InvalidFormatException {
        XWPFTableCell cell = wordConstruct.getCell();
        XWPFRun run = wordConstruct.getRun();
        String text = run.text();
        for (Map.Entry<String, Customization> entry : pictureLabel.entrySet()) {
            Customization customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text)) {
                List<XWPFParagraph> tempParagraphs = cell.getParagraphs();
                for (int j = 0; j < tempParagraphs.size(); j++) {
                    cell.removeParagraph(j);
                }
                XWPFRun newRun = cell.addParagraph().createRun();
                newRun.removeBreak();
                newRun.removeCarriageReturn();
                newRun.addPicture(customization.getPicture(),
                        AnalyzeFileType.getFileType(customization.getPicture()),
                        customization.getPictureName(),
                        Units.toEMU(customization.getWidth()),
                        Units.toEMU(customization.getHeight()));
                cell.removeParagraph(0);
                customization.handle(wordConstruct, index);
                return true;
            }
        }
        return false;
    }

    /**
     * 2019/8/19
     * get the vanish attribute and set the value to STOnOff.FALSE {@link STOnOff#FALSE}
     *
     * @param ctrPr the attribute of the run of the word
     * @author 657518680@qq.com
     * @since 1.0.0
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
     * @since 1.0.0
     */
    private static XWPFParagraph getFirstTableParagraph(XWPFTableCell tableCell) {
        List<XWPFParagraph> paragraphList = tableCell.getParagraphs();
        if (paragraphList.isEmpty()) {
            return tableCell.addParagraph();
        }
        return paragraphList.get(0);
    }

    /**
     * 2019/8/19
     * determine if there is a next row that can be used for table label rather than create a new row
     *
     * @param table    table
     * @param rowIndex rowIndex {@link Index#getRowIndex()}
     * @param style    the string of the style{@linkplain Processor#getTrPrString}
     * @param j        the index of the list of the value of the table label
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    private static boolean isHasNextRow(XWPFTable table, int rowIndex, String style, int j) {
        return table.getRow(rowIndex + 1) != null
                && style.equals(getTrPrString(table.getRow(rowIndex + 1).getCtRow().getTrPr()))
                && j != 0;
    }

    /**
     * 2019/8/19
     * get the string of ctTrPr if null then return empty string
     *
     * @param ctTrPr the attribute of the run of the word
     * @return the string of the ctTrPr
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    private static String getTrPrString(CTTrPr ctTrPr) {
        if (ctTrPr == null) {
            return "";
        }
        return ctTrPr.toString();
    }

}
