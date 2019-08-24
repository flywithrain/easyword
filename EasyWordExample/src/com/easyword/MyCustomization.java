package com.easyword;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.awt.*;
import java.io.InputStream;

/**
 * 2019/8/24 17:10
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class MyCustomization implements Customization {

    private String text;

    public MyCustomization(String s) {
        setText(s);
    }

    @Override
    public void handle(WordConstruct wordConstruct, Index index) {
        XWPFRun run = wordConstruct.getRun();
        run.setBold(true);
        run.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
        run.setFontFamily("Courier");
    }

    @Override
    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    @Override
    public InputStream getPicture() {
        return null;
    }

    @Override
    public String getPictureName() {
        return null;
    }

    @Override
    public int getWidth() {
        return 0;
    }

    @Override
    public int getHeight() {
        return 0;
    }
}
