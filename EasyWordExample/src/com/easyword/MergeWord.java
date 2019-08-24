package com.easyword;

import com.thunisoft.easyword.core.EasyWord;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xmlbeans.XmlException;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * 2019/8/24 13:58
 *
 * @author 657518680@qq.com
 * @since 1.0.0
 */
public class MergeWord {

    public static void main(String[] args) throws IOException, InvalidFormatException, XmlException {
        List<InputStream> wordList = new ArrayList<>(3);
        wordList.add(new FileInputStream(System.getProperty("user.dir") + "\\resources\\staticlabel.docx"));
        wordList.add(new FileInputStream(System.getProperty("user.dir") + "\\resources\\dynamiclabel.docx"));
        wordList.add(new FileInputStream(System.getProperty("user.dir") + "\\resources\\tablelabel.docx"));
        EasyWord.mergeWord(wordList, new FileOutputStream(System.getProperty("user.dir") + "\\result\\mergeword-result.docx"));
    }

}
