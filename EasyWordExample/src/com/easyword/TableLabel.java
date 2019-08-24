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
 * 2019/8/24 13:50
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class TableLabel {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, List<List<Customization>>> tableLabel = new HashMap<>(3);
        List<List<Customization>> rows = new ArrayList<>(3);
        List<Customization> row1 = new ArrayList<>(3);
        row1.add(new DefaultCustomization("11"));
        row1.add(new DefaultCustomization("12"));
        row1.add(new DefaultCustomization("13"));
        rows.add(row1);
        List<Customization> row2 = new ArrayList<>(3);
        row2.add(new DefaultCustomization("21"));
        row2.add(new DefaultCustomization("22"));
        row2.add(new DefaultCustomization("23"));
        rows.add(row2);
        List<Customization> row3 = new ArrayList<>(3);
        row3.add(new DefaultCustomization("31"));
        row3.add(new DefaultCustomization("32"));
        row3.add(new DefaultCustomization("33"));
        rows.add(row3);
        tableLabel.put("row", rows);
        EasyWord.replaceLabel(new FileInputStream(System.getProperty("user.dir") + "\\resources\\tablelabel.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\result\\tablelabel-result.docx"),
                new HashMap<>(0),
                new HashMap<>(0),
                tableLabel,
                new HashMap<>(0));
    }

}
