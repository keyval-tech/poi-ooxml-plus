package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.StringReplace;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class StringReplaceProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        StringReplace stringReplace = (StringReplace) annotation;
        String[] target = stringReplace.target();
        String[] replacement = stringReplace.replacement();

        if (target.length != replacement.length || target.length == 0) {
            throw new PoiOoxmlPlusException("@StringReplace配置有误");
        }

        String strValue = String.valueOf(targetFieldValue);
        for (int i = 0; i < target.length; i++) {
            strValue = strValue
                    .replace(StringUtils.replaceFieldValue(target[i], object),
                            StringUtils.replaceFieldValue(replacement[i], object));
        }
        
        return strValue;
    }
}
