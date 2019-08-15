package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * ProcessTable4Table
 *
 * @author 657518680@qq.com
 * @date 2019/8/14 10:31
 * @since 1.0.0
 */
class ProcessTable4Table {

    private boolean myResult;
    private Map<String, List<List<Customization>>> tableLabel;
    private XWPFTable table;
    private int rowIndex;
    private XWPFTableRow row;
    private XWPFParagraph paragraph;
    private XWPFRun run;
    private String text;

    ProcessTable4Table(Map<String, List<List<Customization>>> tableLabel,
                       XWPFTable table,
                       int rowIndex,
                       XWPFTableRow row,
                       XWPFParagraph paragraph,
                       XWPFRun run,
                       String text) {
        this.tableLabel = tableLabel;
        this.table = table;
        this.rowIndex = rowIndex;
        this.row = row;
        this.paragraph = paragraph;
        this.run = run;
        this.text = text;
    }

    boolean isContinue() {
        return myResult;
    }

    int getRowIndex() {
        return rowIndex;
    }

    void process() {
        for (Map.Entry<String, List<List<Customization>>> entry : tableLabel.entrySet()) {
            String key = entry.getKey();
            List<List<Customization>> listList = entry.getValue();
            if (key.equals(text)) {
                CTTrPr ctTrPr = row.getCtRow().getTrPr();
                String style = getTrPr(ctTrPr);
                List<XWPFTableCell> tableCells = row.getTableCells();
                List<CTTcPr> ctTcPrList = new ArrayList<>();
                for (XWPFTableCell temp : tableCells) {
                    ctTcPrList.add(temp.getCTTc().getTcPr());
                }
                int temp = rowIndex;
                for (int j = 0; j < listList.size(); j++) {
                    List<Customization> list = listList.get(j);
                    XWPFTableRow newTableRow;
                    if (isHasNextRow(style, j)) {
                        newTableRow = table.getRow(rowIndex + 1);
                        for (int k = 0; k < list.size(); k++) {
                            Customization customization = list.get(k);
                            XWPFTableCell tableCell = newTableRow.getCell(k);
                            XWPFParagraph xwpfParagraph = tableCell.getParagraphs().get(0);
                            XWPFRun xwpfRun = xwpfParagraph.createRun();
                            xwpfRun.getCTR().setRPr(run.getCTR().getRPr());
                            xwpfRun.setText(customization.getText());
                        }
                        ++rowIndex;
                    } else {
                        newTableRow = table.insertNewTableRow(rowIndex + 1);
                        for (int k = 0; k < list.size(); k++) {
                            Customization customization = list.get(k);
                            XWPFTableCell newTableCell = newTableRow.addNewTableCell();
                            newTableCell.getCTTc().setTcPr(ctTcPrList.get(k));
                            XWPFParagraph newParagraph = newTableCell.getParagraphs().get(0);
                            newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                            XWPFRun newRun = newParagraph.createRun();
                            newRun.getCTR().setRPr(run.getCTR().getRPr());
                            newRun.setText(customization.getText());
                        }
                        newTableRow.getCtRow().setTrPr(ctTrPr);
                        ++rowIndex;
                    }
                }
                --rowIndex;
                table.removeRow(temp);
                myResult = true;
                return;
            }
        }
        myResult = false;
    }

    private boolean isHasNextRow(String style, int j) {
        return table.getRow(rowIndex + 1) != null
                && style.equals(getTrPr(table.getRow(rowIndex + 1).getCtRow().getTrPr()))
                && j != 0;
    }

    private static String getTrPr(CTTrPr ctTrPr){
        if(ctTrPr == null){
            return "";
        }
        return ctTrPr.toString();
    }

}
