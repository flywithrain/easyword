package com.thunisoft.easyword.bo;

/**
 * 2019/10/26 13:37
 *
 * @author wangxiaoyu 657518680@qq.com
 * @since 1.1.3
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

}
