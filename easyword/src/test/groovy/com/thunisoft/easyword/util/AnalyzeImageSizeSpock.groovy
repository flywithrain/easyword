package com.thunisoft.easyword.util

import spock.lang.Specification

/**
 * @author 65751* @date 2019-08-2019/8/20 14:39
 */
class AnalyzeImageSizeSpock extends Specification{

    def "analyze the type of the file"(){
        when:
        def size = AnalyzeImageSize.getImageSize(this.getClass().getClassLoader().getResourceAsStream("\\file\\zr.jpg"))

        then:
        verifyAll(size) {
            size.get("width") == 1000
            size.get("height") == 1478
        }

    }

}
