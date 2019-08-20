package com.thunisoft.easyword.util

import com.thunisoft.easyword.constant.FileTypeEnum
import org.apache.poi.util.IOUtils
import spock.lang.Specification

/**
 * @author 65751* @date 2019-08-2019/8/15 15:40
 */
class AnalyzeFileTypeSpock extends Specification {

    def "analyze the type of the file"() {
        expect:
        AnalyzeFileType.getFileType(IOUtils.toByteArray(this.getClass().getClassLoader().getResourceAsStream('\\file\\' + file))) == type

        where:
        file     | type
        '1.docx' | -1
        '1.jpg'  | FileTypeEnum.JPEG.value
        '1.pdf'  | -1
        '1.png'  | FileTypeEnum.PNG.value
    }

}
