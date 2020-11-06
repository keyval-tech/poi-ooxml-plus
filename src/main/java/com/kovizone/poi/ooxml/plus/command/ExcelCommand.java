package com.kovizone.poi.ooxml.plus.command;

import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.util.ElParser;
import org.apache.poi.ss.usermodel.*;

import java.util.*;

/**
 * 工作表基础命令
 *
 * @author KoviChen
 */
public class ExcelCommand extends WriteSheetCommand {

    /**
     * EL解析器
     */
    private ElParser elParser;
    /**
     * 数据遍历索引
     */
    private Integer entityListIndex;
    /**
     * 表达式解析器
     */
    private List<?> entityList;

    public ExcelCommand(
            Workbook workbook,
            ExcelStyle excelStyle,
            int cellSize,
            Map<String, Object> replaceMap,
            List<?> entityList,
            ExcelInitCommand excelInitCommand) {
        super();

        this.elParser = new ElParser(entityList);
        this.entityList = entityList;
        this.entityListIndex = 0;

        super.setCellSize(cellSize);

        super.setWorkbook(workbook);
        super.setReplaceMap(replaceMap);

        super.setDefaultColumnWidth(excelInitCommand.getDefaultColumnWidth());
        super.setDefaultRowHeight(excelInitCommand.getDefaultRowHeight());
        super.setBlankRowHeights(excelInitCommand.getBlankRowHeights());
        super.setBlankCellWidths(excelInitCommand.getBlankCellWidths());
        super.setDefaultSheetName(excelInitCommand.getDefaultSheetName());

        super.setStyleMap(excelStyle.styleMap(new ExcelStyleCommand(workbook)));

        super.setLazyRenderCellWidth(new HashMap<>(16));
        super.setLazyRenderRowHeight(new HashMap<>(16));

    }

    public <T> T parse(String expression, Class<T> clazz) {
        return elParser.parse(expression, entityListIndex, clazz);
    }

    public String parseString(String expression) {
        return elParser.parseString(expression, entityListIndex);
    }

    public Boolean parseBoolean(String expression) {
        return elParser.parseBoolean(expression, entityListIndex);
    }

    public List<?> getEntityList() {
        return entityList;
    }

    public int currentEntityListIndex() {
        return entityListIndex;
    }

    public void nextEntityListIndex() {
        if (entityListIndex == null) {
            entityListIndex = 0;
        } else {
            entityListIndex++;
        }
    }
}
