package com.thunisoft.easyword.bo;

import org.apache.poi.xwpf.usermodel.*;

import java.io.InputStream;

/**
 * Customization
 *
 * @author 657518680@qq.com
 * @date 2019/8/13 10:50
 * @since 1.0.0
 */
public interface Customization {

    default void handle(WordConstruct wordConstruct, Index index) {
        // do nothing if need can override
    }

    String getText();

    InputStream getPicture();

    String getPictureName();

    int getWidth();

    int getHeight();

}
