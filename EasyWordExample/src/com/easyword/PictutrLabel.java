package com.easyword;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.DefaultCustomization;
import com.thunisoft.easyword.core.EasyWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 2019/8/24 13:40
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class PictutrLabel {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, Customization> pictureLabel = new HashMap<>(0);
        DefaultCustomization defaultCustomization = new DefaultCustomization();
        defaultCustomization.setPicture(new FileInputStream(System.getProperty("user.dir") + "\\resources\\zrqk.jpg"));
        defaultCustomization.setPictureName("昨日青空");
        pictureLabel.put("zrqk", defaultCustomization);
        EasyWord.replaceLabel(new FileInputStream(System.getProperty("user.dir") + "\\resources\\picturelabel.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\result\\picturelabel-result.docx"),
                new HashMap<>(0),
                new HashMap<>(0),
                new HashMap<>(0),
                pictureLabel);
    }

}
