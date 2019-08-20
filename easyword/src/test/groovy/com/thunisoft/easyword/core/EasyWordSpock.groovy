package com.thunisoft.easyword.core

import com.thunisoft.easyword.bo.Customization
import com.thunisoft.easyword.bo.DefaultCustomization
import spock.lang.Specification
import spock.lang.Stepwise

import java.time.LocalDateTime

/**
 * @author 65751* @date 2019-08-2019/8/15 18:55
 */
@Stepwise
class EasyWordSpock extends Specification {

    def "test replace label correct"() {
        given:
        Map<String, Customization> staticLabel =
                ["tjsj": new DefaultCustomization(LocalDateTime.now().toString()),
                 "qm"  : new DefaultCustomization("EasyWord-Spock")]
        Map<String, List<Customization>> dynamicLabel = ['bc': [new DefaultCustomization('Programming Language  Ratings'),
                                                                new DefaultCustomization('Java\t16.028%'),
                                                                new DefaultCustomization('C\t15.154%'),
                                                                new DefaultCustomization('Python\t10.020%'),
                                                                new DefaultCustomization('C++\t6.057%'),
                                                                new DefaultCustomization('C#\t3.842%'),
                                                                new DefaultCustomization('Visual Basic .NET\t3.695%'),
                                                                new DefaultCustomization('JavaScript\t2.258%'),
                                                                new DefaultCustomization('PHP\t2.075%'),
                                                                new DefaultCustomization('Objective-C\t1.690%'),
                                                                new DefaultCustomization('SQL\t1.625%\t-0.69%')]]
        List<List<Customization>> lists = [[new DefaultCustomization('1'), new DefaultCustomization("战狼2"), new DefaultCustomization("56.39"), new DefaultCustomization("2017")],
                                           [new DefaultCustomization("2"), new DefaultCustomization("流浪地球"), new DefaultCustomization("46.18"), new DefaultCustomization('2019')],
                                           [new DefaultCustomization("3"), new DefaultCustomization("复仇者联盟4：终局之战"), new DefaultCustomization("42.05"), new DefaultCustomization('2019')],
                                           [new DefaultCustomization("4"), new DefaultCustomization("哪吒之魔童降世"), new DefaultCustomization("41.35"), new DefaultCustomization('2019')],
                                           [new DefaultCustomization("5"), new DefaultCustomization("红海行动"), new DefaultCustomization("36.22"), new DefaultCustomization('2018')],
                                           [new DefaultCustomization("6"), new DefaultCustomization("美人鱼"), new DefaultCustomization("33.9"), new DefaultCustomization('2016')],
                                           [new DefaultCustomization("7"), new DefaultCustomization("唐人街探案2"), new DefaultCustomization("33.71"), new DefaultCustomization('2018')],
                                           [new DefaultCustomization("8"), new DefaultCustomization("我不是药神"), new DefaultCustomization("30.75"), new DefaultCustomization('2018')],
                                           [new DefaultCustomization("9"), new DefaultCustomization("速度与激情8"), new DefaultCustomization("26.49"), new DefaultCustomization('2017')],
                                           [new DefaultCustomization("10"), new DefaultCustomization("西虹市首富"), new DefaultCustomization("25.27"), new DefaultCustomization('2018')]]
        Map<String, List<List<Customization>>> tableLabel = ["dy": lists]
        def picture = new DefaultCustomization()
        picture.setPicture(this.getClass().getClassLoader().getResourceAsStream("\\file\\zr.jpg"))
        picture.setPictureName('哪吒之魔童降世')
        Map<String, Customization> pictureLabel = ["zr": picture]
        EasyWord.replaceLabel(this.getClass().getClassLoader().getResourceAsStream("\\file\\1.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\replace.docx"),
                staticLabel,
                dynamicLabel,
                tableLabel,
                pictureLabel)

        expect:
        true
    }

    def "test merge word correct"() {
        given:
        EasyWord.mergeWord([this.getClass().getClassLoader().getResourceAsStream("\\file\\1.docx"),
                            new FileInputStream(System.getProperty("user.dir") + "\\replace.docx")],
                new FileOutputStream(System.getProperty("user.dir") + "\\merge.docx"))

        expect:
        true
    }

}
