package com.easyword;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.DefaultCustomization;
import com.thunisoft.easyword.core.EasyWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 2019/8/24 17:08
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class CustomizationDemo {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, List<Customization>> dynamicLabel = new HashMap<>(1);
        List<Customization> list = new ArrayList<>(10);
        list.add(new MyCustomization("Programming Language  Ratings"));
        list.add(new MyCustomization("Java\t16.028%"));
        list.add(new MyCustomization("C\t15.154%"));
        list.add(new MyCustomization("Python\t10.020%"));
        list.add(new MyCustomization("C++\t6.057%"));
        list.add(new MyCustomization("C#\t3.842%"));
        list.add(new MyCustomization("Visual Basic .NET\t3.695%"));
        list.add(new MyCustomization("JavaScript\t2.258%"));
        list.add(new MyCustomization("PHP\t2.075%"));
        list.add(new MyCustomization("Objective-C\t1.690%"));
        dynamicLabel.put("bc", list);
        EasyWord.replaceLabel(new FileInputStream(System.getProperty("user.dir") + "\\resources\\dynamiclabel.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\result\\customizationDemo-result.docx"),
                new HashMap<>(0),
                dynamicLabel,
                new HashMap<>(0),
                new HashMap<>(0));
    }

}
