package com.thunisoft.easyword.core

import com.thunisoft.easyword.bo.Customization
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
        def staticLabelite = ["tjsj": LocalDateTime.now().toString(),
                              "qm"  : "EasyWord-Spock"]
        def dynamicLabelite = ['bc': ['Programming Language  Ratings',
                                      'Java\t16.028%',
                                      'C\t15.154%',
                                      'Python\t10.020%',
                                      'C++\t6.057%',
                                      'C#\t3.842%',
                                      'Visual Basic .NET\t3.695%',
                                      'JavaScript\t2.258%',
                                      'PHP\t2.075%',
                                      'Objective-C\t1.690%',
                                      'SQL\t1.625%\t-0.69%']]
        def lists = [['1', "战狼2", "56.39", "2017"],
                     ["2", "流浪地球", "46.18", '2019'],
                     ["3", "复仇者联盟4：终局之战", "42.05", '2019'],
                     ["4", "哪吒之魔童降世", "41.35", '2019'],
                     ["5", "红海行动", "36.22", '2018'],
                     ["6", "美人鱼", "33.9", '2016'],
                     ["7", "唐人街探案2", "33.71", '2018'],
                     ["8", "我不是药神", "30.75", '2018'],
                     ["9", "速度与激情8", "26.49", '2017'],
                     ["10", "西虹市首富", "25.27", '2018']]
        def tableLabelite = ["dy": lists]
        def verticalLabelite = ['vb' : ['1', '2', '3', '4', '5', '6'],
                                'vb1': ['1', '2', '3', '4', '5', '6']]

        def staticLabel = StaticLabelImp.lite2Full(staticLabelite)
        def dynamicLabel = DynamicLabelImp.lite2Full(dynamicLabelite)
        def tableLabel = TabelLabelImp.lite2Full(tableLabelite)
        def verticalLabel = VerticalLabelImp.lite2Full(verticalLabelite)
        ((VerticalLabelImp)verticalLabel.get('vb1')).setRowSum(6)
        def picture = new PictureLabelImp()
        picture.setPicture(this.getClass().getClassLoader().getResourceAsStream("\\file\\zr.jpg"))
        picture.setPictureName('哪吒之魔童降世')
        def pictureLabel = ["zr": picture]

        staticLabel.putAll(dynamicLabel)
        staticLabel.putAll(tableLabel)
        staticLabel.putAll(verticalLabel)
        staticLabel.putAll(pictureLabel)

        EasyWord.replaceLabel(this.getClass().getClassLoader().getResourceAsStream("\\file\\1.docx"),
                new FileOutputStream(System.getProperty("user.dir") + "\\replace.docx"),
                staticLabel)

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
