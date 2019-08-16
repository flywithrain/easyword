package com.thunisoft.easyword.core

import com.thunisoft.easyword.bo.Customization
import com.thunisoft.easyword.bo.DefaultCustomization
import spock.lang.Specification

/**
 * @author 65751* @date 2019-08-2019/8/15 18:55
 */
class EasyWordSpock extends Specification{

    def "test correct"(){
        given:
        Map<String, Customization> staticLabel = new HashMap<>(1);
        staticLabel.put("grc", new DefaultCustomization("顾涌泉"));
        Map<String, List<List<Customization>>> tableLabel = new HashMap<>(1);
        List<List<Customization>> lists = new ArrayList<>(2);
        List<Customization> list1 = new ArrayList<>(3);
        list1.add(new DefaultCustomization("顾涌泉11"));
        list1.add(new DefaultCustomization("顾涌泉12"));
        list1.add(new DefaultCustomization("顾涌泉13"));
        lists.add(list1);
        List<Customization> list2 = new ArrayList<>(3);
        list2.add(new DefaultCustomization("顾涌泉21"));
        list2.add(new DefaultCustomization("顾涌泉22"));
        list2.add(new DefaultCustomization("顾涌泉23"));
        lists.add(list2);
        tableLabel.put("gyq", lists);
        EasyWord.replaceLabel(this.getClass().getClassLoader().getResourceAsStream("\\file\\1.docx"),
                new ByteArrayOutputStream(1024),
                staticLabel,
                new HashMap<>(0),
                tableLabel,
                new HashMap<>(0))

        expect:
        true
    }

}
