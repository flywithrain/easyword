package com.easyword;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.DefaultCustomization;
import com.thunisoft.easyword.core.EasyWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * 2019/8/23 18:08
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class StaticLabel {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, Customization> staticLabel = new HashMap<>(3);
        staticLabel.put("dy", new DefaultCustomization("回填会默认使用表格默认格式"));
        staticLabel.put("tjsj", new DefaultCustomization("段落部分的回填也会使用段落默认格式，即标签格式"));
        EasyWord.replaceLabel(new FileInputStream(System.getProperty("user.dir") + "\\resources\\staticlabel.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\result\\staticlabel-result.docx"),
                staticLabel);
    }

}
