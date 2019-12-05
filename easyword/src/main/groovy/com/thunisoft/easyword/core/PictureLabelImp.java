package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import com.thunisoft.easyword.util.AnalyzeFileType;
import com.thunisoft.easyword.util.AnalyzeImageSize;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.thunisoft.easyword.core.Processor.processVanish;

/**
 * 2019/12/3 17:57
 *
 * @author wangxiaoyu 657518680@qq.com
 * @since 2.0.0
 */
public class PictureLabelImp implements Customization {

    Logger logger = Logger.getLogger("EasyWordLogger");

    private InputStream picture;
    private String pictureName;
    private int width;
    private int height;

    public PictureLabelImp() {
    }

    public InputStream getPicture() {
        return picture;
    }

    public String getPictureName() {
        return pictureName;
    }

    public int getWidth() {
        return width;
    }

    public int getHeight() {
        return height;
    }

    public void setPicture(InputStream picture) {
        this.picture = picture;
    }

    public void setPictureName(String pictureName) {
        this.pictureName = pictureName;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public void setHeight(int height) {
        this.height = height;
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
        XWPFParagraph paragraph = wordConstruct.getParagraph();
        XWPFRun run = wordConstruct.getRun();
        int rIndex = index.getrIndex();
        XWPFRun newRun = paragraph.insertNewRun(rIndex);
        CTRPr ctrPr = run.getCTR().getRPr();
        processVanish(ctrPr);
        newRun.getCTR().setRPr(ctrPr);
        paragraph.removeRun(rIndex + 1);
        processPicture(newRun);
    }

    /**
     * 2019/8/20 14:15
     *
     * @param newRun the run created to save image
     * @author 657518680@qq.com
     * @since beta
     */
    private void processPicture(XWPFRun newRun) {
        try {
            byte[] bytes = IOUtils.toByteArray(picture);
            if (width <= 0 || height <= 0) {
                Map<String, Integer> size = AnalyzeImageSize.getImageSize(new ByteArrayInputStream(bytes));
                width = size.get("width");
                height = size.get("height");
            }
            newRun.addPicture(new ByteArrayInputStream(bytes),
                    AnalyzeFileType.getFileType(bytes),
                    pictureName,
                    Units.pixelToEMU(width),
                    Units.pixelToEMU(height));
        } catch (Exception e) {
            logger.log(Level.SEVERE, "PictureLabelImp: processPicture failed!", e);
        }
    }

}
