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
 * 2019/12/4 16:12
 *
 * @author wangxiaoyu 657518680@qq.com
 * @version 2.0.0
 * @since 2.0.0
 */
public class VerticalLabelImp implements Customization {

    private static final Logger logger = Logger.getLogger("EasyWordLogger");

    private List<String> list;

    private int rowSum;

    public VerticalLabelImp() {
    }

    public VerticalLabelImp(List<String> list) {
        setList(list);
    }

    public VerticalLabelImp(List<String> list, int rowSum) {
        setList(list);
        this.rowSum = rowSum;
    }

    public List<String> getList() {
        return list;
    }

    public void setList(List<String> list) {
        if (list == null) {
            this.list = new ArrayList<>(0);
        } else {
            this.list = list;
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
    public void handle(String key, WordConstruct wordConstruct, Index index) {
        if (CollectionUtils.isEmpty(list)) {
            return;
        }

        XWPFTable table = wordConstruct.getTable();
        XWPFTableRow row = wordConstruct.getRow();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();

        int rowIndex = index.getRowIndex();
        int cellIndex = index.getcIndex();

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
            logger.log(Level.SEVERE, "VerticalLabelImp：deepClone failed", e);
        }

        List<XWPFTableCell> tableCells = row.getTableCells();
        List<CTTcPr> ctTcPrList = new ArrayList<>();
        for (XWPFTableCell temp : tableCells) {
            ctTcPrList.add(temp.getCTTc().getTcPr());
        }

        for (int i = 0; i < list.size(); i++) {
            String str = list.get(i);
            if (isTheNextRow(table, rowIndex, style, i)) {
                XWPFTableRow tempRow = table.getRow(rowIndex);
                XWPFTableCell tempCell = tempRow.getCell(cellIndex);
                clearCell(tempCell);
                XWPFParagraph tempParagraph = getFirstTableParagraph(tempCell);
                tempParagraph.getCTP().setPPr(deepCopyPpr);
                XWPFRun tempRun = tempParagraph.createRun();
                tempRun.getCTR().setRPr(deepCopyRpr);
                tempRun.setText(str);
            } else {
                XWPFTableRow newTableRow = table.insertNewTableRow(rowIndex);
                newTableRow.getCtRow().setTrPr(ctTrPr);
                for (int k = 0; k < tableCells.size(); k++) {
                    XWPFTableCell newCell = newTableRow.addNewTableCell();
                    newCell.getCTTc().setTcPr(ctTcPrList.get(k));
                    if (k == cellIndex) {
                        XWPFParagraph newParagraph = getFirstTableParagraph(newCell);
                        newParagraph.getCTP().setPPr(deepCopyPpr);
                        XWPFRun newRun = newParagraph.createRun();
                        newRun.getCTR().setRPr(deepCopyRpr);
                        newRun.setText(str);
                    }
                }
            }
            rowIndex++;
        }
        index.setpIndex(1);
    }

    public boolean isTheNextRow(XWPFTable table, int rowIndex, String style, int j) {
        if (rowSum > 0) {
            return rowSum > j;
        }
        return Processor.isTheNextRow(table, rowIndex, style, j);
    }

    public static Map<String, Customization> lite2Full(Map<String, List<String>> verticalLabelite) {
        return verticalLabelite.entrySet().stream()
                .collect(Collectors.toMap(Map.Entry::getKey, entry -> new VerticalLabelImp(entry.getValue())));
    }

}
