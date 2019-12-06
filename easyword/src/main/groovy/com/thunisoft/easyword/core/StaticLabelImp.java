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
 * @version 2.0.0
 */
public class StaticLabelImp implements Customization {

    private String text;

    public StaticLabelImp() {
        setText("");
    }

    public StaticLabelImp(String text) {
        setText(text);
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        if (text == null) {
            this.text = "";
        } else {
            this.text = text;
        }
    }

    /**
     * 2019/12/6
     * static label back fill
     *
     * @param wordConstruct the struct of word in POI in paragraph only paragraph and run available
     * @param index         the index of attributes in wordConstruct
     * @param key           the label
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
     * 2019/12/6 14:48
     * Convert staticLabelite to staticLabel
     *
     * @param staticLabelite a simplified version of staticLabel
     * @return staticLabel
     * @author 657518680@qq.com
     * @since 2.0.0
     */
    public static Map<String, Customization> lite2Full(Map<String, String> staticLabelite) {
        return staticLabelite.entrySet().stream()
                .collect(Collectors.toMap(Map.Entry::getKey, entry -> new StaticLabelImp(entry.getValue())));
    }

}
