package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTPPrImpl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTRPrImpl;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Collectors;

import static com.thunisoft.easyword.core.Processor.*;

/**
 * 2019/12/4 10:07
 *
 * @author wangxiaoyu 657518680@qq.com
 * @since 2.0.0
 */
public class TabelLabelImp implements Customization {

    private static final Logger logger = Logger.getLogger("EasyWordLogger");

    private List<List<String>> tableList;

    private int rowSum;

    public TabelLabelImp() {
    }

    public TabelLabelImp(List<List<String>> tableList) {
        setTableList(tableList);
    }

    public TabelLabelImp(List<List<String>> tableList, int rowSum) {
        setTableList(tableList);
        this.rowSum = rowSum;
    }

    public List<List<String>> getTableList() {
        return tableList;
    }

    public void setTableList(List<List<String>> tableList) {
        if (tableList == null) {
            this.tableList = new ArrayList<>(0);
        } else {
            this.tableList = tableList;
        }
    }

    public int getRowSum() {
        return rowSum;
    }

    public void setRowSum(int rowSum) {
        this.rowSum = rowSum;
    }

    /**
     * 2019/8/19
     * By implementing this method you can do almost anything with word
     *
     * @param wordConstruct the struct of word in POI in paragraph only paragraph and run available
     * @param index         the index of attributes in wordConstruct
     * @author 657518680@qq.com
     * @since alpha
     */
    @Override
    public void handle(WordConstruct wordConstruct, Index index) {

        if (CollectionUtils.isEmpty(tableList)) {
            return;
        }

        XWPFTable table = wordConstruct.getTable();
        XWPFTableRow row = wordConstruct.getRow();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();

        int rowIndex = index.getRowIndex();
        int cIndexMax = row.getTableCells().size();
        int rIndexMax = paragraph.getRuns().size();

        CTTrPr ctTrPr = row.getCtRow().getTrPr();
        String style = getTrPrString(ctTrPr);
        CTPPr ctpPr = paragraph.getCTP().getPPr();
        CTRPr ctrPr = run.getCTR().getRPr();
        CTPPr deepCopyPpr = null;
        CTRPr deepCopyRpr = null;
        try {
            deepCopyPpr = deepClone((CTPPrImpl) ctpPr);
            deepCopyRpr = deepClone((CTRPrImpl) ctrPr);
            processVanish(deepCopyRpr);
        } catch (IOException | ClassNotFoundException e) {
            logger.log(Level.SEVERE, "TabelLabelImp：deepClone failed", e);
        }

        List<XWPFTableCell> tableCells = row.getTableCells();
        List<CTTcPr> ctTcPrList = new ArrayList<>();
        for (XWPFTableCell temp : tableCells) {
            ctTcPrList.add(temp.getCTTc().getTcPr());
        }

        for (int j = 0; j < tableList.size(); j++) {
            List<String> list = tableList.get(j);
            if (isTheNextRow(table, rowIndex, style, j)) {
                XWPFTableRow newTableRow = table.getRow(rowIndex);
                for (int k = 0; k < list.size(); k++) {
                    String str = list.get(k);
                    XWPFTableCell tableCell = newTableRow.getCell(k);
                    clearCell(tableCell);
                    XWPFParagraph xwpfParagraph = getFirstTableParagraph(tableCell);
                    xwpfParagraph.getCTP().setPPr(deepCopyPpr);
                    XWPFRun xwpfRun = xwpfParagraph.createRun();
                    xwpfRun.getCTR().setRPr(deepCopyRpr);
                    xwpfRun.setText(str);
                }
            } else {
                XWPFTableRow newTableRow = table.insertNewTableRow(rowIndex);
                newTableRow.getCtRow().setTrPr(ctTrPr);
                for (int k = 0; k < list.size(); k++) {
                    XWPFTableCell newTableCell = newTableRow.addNewTableCell();
                    String str = list.get(k);
                    if (k < ctTcPrList.size()) {
                        newTableCell.getCTTc().setTcPr(ctTcPrList.get(k));
                    } else {
                        newTableCell.getCTTc().setTcPr(ctTcPrList.get(ctTcPrList.size() - 1));
                    }
                    XWPFParagraph newParagraph = getFirstTableParagraph(newTableCell);
                    newParagraph.getCTP().setPPr(deepCopyPpr);
                    XWPFRun newRun = newParagraph.createRun();
                    newRun.getCTR().setRPr(deepCopyRpr);
                    newRun.setText(str);
                }
            }
            ++rowIndex;
        }
        rowIndex--;
        index.setrIndex(rIndexMax);
        index.setpIndex(1);
        index.setcIndex(cIndexMax);
        index.setRowIndex(rowIndex);
    }

    public boolean isTheNextRow(XWPFTable table, int rowIndex, String style, int j) {
        if (rowSum > 0) {
            return rowSum > j;
        }
        return Processor.isTheNextRow(table, rowIndex, style, j);
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
    public static Map<String, Customization> lite2Full(Map<String, List<List<String>>> tableLabelite) {
        return tableLabelite.entrySet().stream()
                .collect(Collectors.toMap(Map.Entry::getKey, entry -> new TabelLabelImp(entry.getValue())));
    }

}
