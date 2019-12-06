package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 2019/12/3 19:59
 *
 * @author wangxiaoyu 657518680@qq.com
 * @version 2.0.0
 * @since 2.0.0
 */
public class DynamicLabelImp implements Customization {

    private List<String> list;

    public DynamicLabelImp() {
    }

    public DynamicLabelImp(List<String> list) {
        setList(list);
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

    /**
     * 2019/12/6
     * dynamic label back fill
     *
     * @param wordConstruct the struct of word in POI in paragraph only paragraph and run available
     * @param index         the index of attributes in wordConstruct
     * @param key           the label
     * @author 657518680@qq.com
     * @since 2.0.0
     */
    @Override
    public void handle(String key, WordConstruct wordConstruct, Index index) {
        XWPFDocument document = wordConstruct.getDocument();
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int pIndex = index.getpIndex();
        for (String str : list) {
            XmlCursor cursor = paragraph.getCTP().newCursor();
            XWPFParagraph newPara = document.insertNewParagraph(cursor);
            newPara.getCTP().setPPr(paragraph.getCTP().getPPr());
            XWPFRun newRun = newPara.createRun();
            newRun.getCTR().setRPr(run.getCTR().getRPr());
            newRun.setText(str);
        }
        document.removeBodyElement(document.getPosOfParagraph(paragraph));
        index.setpIndex(pIndex + list.size() - 1);
    }

    /**
     * 2019/12/6 14:48
     * Convert dynamicLabelite to dynamicLabel
     *
     * @param dynamicLabelite a simplified version of dynamicLabel
     * @return dynamicLabel
     * @author 657518680@qq.com
     * @since 2.0.0
     */
    public static Map<String, Customization> lite2Full(Map<String, List<String>> dynamicLabelite) {
        return dynamicLabelite.entrySet().stream()
                .collect(Collectors.toMap(Map.Entry::getKey, entry -> new DynamicLabelImp(entry.getValue())));
    }

}
