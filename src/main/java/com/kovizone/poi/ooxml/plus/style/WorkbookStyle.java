package com.kovizone.poi.ooxml.plus.style;

import com.kovizone.poi.ooxml.plus.WorkbookStyleCommand;
import org.apache.poi.ss.usermodel.*;

public interface WorkbookStyle {

    /**
     * 表头标题行样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle headerTitleRowStyle(WorkbookStyleCommand command);

    /**
     * 表头标题样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle headerTitleCellStyle(WorkbookStyleCommand command);

    /**
     * 表头副标题行样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle headerSubtitleRowStyle(WorkbookStyleCommand command);

    /**
     * 表头副标题样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle headerSubtitleCellStyle(WorkbookStyleCommand command);

    /**
     * 数据列名行样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle dateTitleRowStyle(WorkbookStyleCommand command);

    /**
     * 数据列名样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle dateTitleCellStyle(WorkbookStyleCommand command);

    /**
     * 数据内容行样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle dateBodyRowStyle(WorkbookStyleCommand command);

    /**
     * 数据内容样式
     *
     * @param command 命令
     * @return CellStyle 样式
     */
    CellStyle dateBodyCellStyle(WorkbookStyleCommand command);
}
