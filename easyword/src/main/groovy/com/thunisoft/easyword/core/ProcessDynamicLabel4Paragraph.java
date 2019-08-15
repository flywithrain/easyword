package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.util.List;
import java.util.Map;

/**
 * ProcessDynamicLabel4Paragraph
 *
 * @author 657518680@qq.com
 * @date 2019/8/13 19:39
 * @since 1.0.0
 */
class ProcessDynamicLabel4Paragraph {

    private boolean myResult;
    private XWPFDocument xwpfDocument;
    private Map<String, List<Customization>> dynamicLabel;
    private int pIndex;
    private XWPFParagraph paragraph;
    private XWPFRun run;
    private String text;

    ProcessDynamicLabel4Paragraph(XWPFDocument xwpfDocument,
                                         Map<String, List<Customization>> dynamicLabel,
                                         int pIndex,
                                         XWPFParagraph paragraph,
                                         XWPFRun run,
                                         String text) {
        this.xwpfDocument = xwpfDocument;
        this.dynamicLabel = dynamicLabel;
        this.pIndex = pIndex;
        this.paragraph = paragraph;
        this.run = run;
        this.text = text;
    }

    boolean isContinue(){
        return myResult;
    }

    int getpIndex() {
        return pIndex;
    }

    void process(){
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
                }
                xwpfDocument.removeBodyElement(xwpfDocument.getPosOfParagraph(paragraph));
                pIndex += customizationList.size();
                myResult = true;
                return;
            }
        }
        myResult = false;
    }

}
