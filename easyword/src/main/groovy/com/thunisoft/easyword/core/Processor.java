package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.util.AnalyzeFileType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Processor
 *
 * @author 657518680@qq.com
 * @date 2019/8/13 19:07
 * @since 1.0.0
 */
final class Processor {

    private Processor() {
    }

    static void processStaticLabel(Map<String, Customization> staticLabel,
                                   XWPFParagraph paragraph,
                                   int rIndex,
                                   XWPFRun run,
                                   String text) {
        for (Map.Entry<String, Customization> entry : staticLabel.entrySet()) {
            Customization customization = entry.getValue();
            String key = entry.getKey();
            if (key.equals(text.trim())) {
                XWPFRun newRun = paragraph.insertNewRun(rIndex);
                CTRPr ctrPr = run.getCTR().getRPr();
                processVanish(ctrPr);
                newRun.getCTR().setRPr(ctrPr);
                paragraph.removeRun(rIndex + 1);
                newRun.setText(text.replaceAll(key, customization.getText()));
            }
        }
    }

    static void processPicture4Paragraph(Map<String, Customization> pictureLabel,
                                         XWPFParagraph paragraph,
                                         int rIndex,
                                         XWPFRun run,
                                         String text) throws IOException, InvalidFormatException {
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
            }
        }
    }

    static void processPicture4Table(Map<String, Customization> pictureLabel,
                                     XWPFTableCell cell,
                                     String text) throws IOException, InvalidFormatException {
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
            }
        }
    }

    static void processVanish(CTRPr ctrPr) {
        if (ctrPr != null) {
            CTOnOff vanish = ctrPr.getVanish();
            if (vanish != null && !vanish.isSetVal()) {
                vanish.setVal(STOnOff.FALSE);
            }
        }
    }

}
