package com.kovizone.poi.ooxml.plus.style;

import com.kovizone.poi.ooxml.plus.WorkbookStyleCommand;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/**
 * 样式管理器接口
 *
 * @author KoviChen
 */
public interface WorkbookStyleManager {

    /**
     * 单元格样式管理器
     *
     * @param command 样式创建指令
     * @return 样式Map
     */
    Map<String, CellStyle> styleMap(WorkbookStyleCommand command);
}
