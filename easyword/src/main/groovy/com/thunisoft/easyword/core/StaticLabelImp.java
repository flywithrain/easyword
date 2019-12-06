package com.thunisoft.easyword.core;

import com.thunisoft.easyword.bo.Customization;
import com.thunisoft.easyword.bo.Index;
import com.thunisoft.easyword.bo.WordConstruct;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;

import java.util.Map;
import java.util.stream.Collectors;

import static com.thunisoft.easyword.core.Processor.clearRun;
import static com.thunisoft.easyword.core.Processor.processVanish;

/**
 * 2019/12/3 16:54
 *
 * @author wangxiaoyu 657518680@qq.com
 * @since 2.0.0
 */
public class StaticLabelImp implements Customization {

    private String text;

    public StaticLabelImp() {
        setText("");
    }

    /**
     * 2019/12/3 17:04
     *
     * @param text A String
     * @author wangxiaoyu 657518680@qq.com
     * @since 2.0.0
     */
    public StaticLabelImp(String text) {
        setText(text);
    }

    /**
     * 2019/8/19
     * get the text of the label
     *
     * @return the text of the label
     * @author 657518680@qq.com
     * @since 2.0.0
     */
    public String getText() {
        return text;
    }

    /**
     * 2019/12/3 17:02
     *
     * @param text A String
     * @author wangxiaoyu 657518680@qq.com
     * @since 2.0.0
     */
    public void setText(String text) {
        if (text == null) {
            this.text = "";
        } else {
            this.text = text;
        }
    }

    /**
     * 2019/8/19
     * By implementing this method you can do almost anything with word
     *
     * @param wordConstruct the struct of word in POI in paragraph only paragraph and run available
     * @param index         the index of attributes in wordConstruct
     * @author 657518680@qq.com
     * @since 2.0.0
     */
    @Override
    public void handle(String key, WordConstruct wordConstruct, Index index) {
        XWPFRun run = wordConstruct.getRun();
        CTRPr ctrPr = run.getCTR().getRPr();
        processVanish(ctrPr);
        String str = run.text().replace(key, text);
        clearRun(run).setText(str);
    }

    /**
     * 2019/8/24 14:48
     * Convert staticLabelite to staticLabel
     *
     * @param staticLabelite a simplified version of staticLabel
     * @return staticLabel
     * @author 657518680@qq.com
     * @since 1.0.0
     */
    public static Map<String, Customization> lite2Full(Map<String, String> staticLabelite) {
        return staticLabelite.entrySet().stream()
                .collect(Collectors.toMap(Map.Entry::getKey, entry -> new StaticLabelImp(entry.getValue())));
    }

}
