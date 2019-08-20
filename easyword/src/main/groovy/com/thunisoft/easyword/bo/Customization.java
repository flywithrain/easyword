package com.thunisoft.easyword.bo;

import java.io.InputStream;

/**
 * 2019/8/13 10:50
 * implement the interface to customize your requirement
 *
 * @author 657518680@qq.com
 * @since alpha
 * @version beta
 */
public interface Customization {

    /**
     * 2019/8/19
     * By implementing this method you can do almost anything with word
     *
     * @param wordConstruct the struct of word in POI in paragraph only paragraph and run available
     * @param index         the index of attributes in wordConstruct
     * @author 657518680@qq.com
     * @since alpha
     */
    default void handle(WordConstruct wordConstruct, Index index) {
        // do nothing if need can override
    }

    /**
     * 2019/8/19
     * get the text of the label
     *
     * @return the text of the label
     * @author 657518680@qq.com
     * @since alpha
     */
    String getText();

    /**
     * 2019/8/19
     * get the inputStream of the picture in EasyWord
     *
     * @return the inputStream of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    InputStream getPicture();

    /**
     * 2019/8/19
     * get the name of the picture inputStream in EasyWord
     *
     * @return the name of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    String getPictureName();

    /**
     * 2019/8/19
     * get the width in pixel of the picture inputStream in EasyWord
     * if <=0 will use the native size of the image both height {@link Customization#getHeight()}
     *
     * @return the width of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    int getWidth();

    /**
     * 2019/8/19
     * get the height in pixel of the picture inputStream in EasyWord
     * if <=0 will use the native size of the image both width {@link Customization#getWidth()}
     *
     * @return the height of the picture
     * @author 657518680@qq.com
     * @since alpha
     */
    int getHeight();

}
