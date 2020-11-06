package com.kovizone.poi.ooxml.plus.command;

import com.kovizone.poi.ooxml.plus.ExcelWriter;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class WriteSheetCommand extends WriteRowCommand {

    /**
     * 默认Sheet标签名
     */
    private String defaultSheetName;
    /**
     * 工作簿索引，当工作表不存在工作簿时为-1
     */
    private int sheetIndex;

    public WriteSheetCommand() {
        sheetIndex = -1;
    }

    /**
     * 获取当前工作簿索引
     *
     * @return 当前工作簿索引
     */
    public int currentSheetIndex() {
        return sheetIndex;
    }

    /**
     * 创建工作簿
     */
    public void createSheet() {
        createSheet(this.defaultSheetName);
    }

    /**
     * 创建工作簿
     *
     * @param sheetName 工作簿名
     */
    public void createSheet(String sheetName) {
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = this.defaultSheetName;
        }
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = "Sheet" + ExcelWriter.SHEET_NUM;
        }
        if (!sheetName.contains(ExcelWriter.SHEET_NUM)) {
            sheetName = sheetName + "_" + ExcelWriter.SHEET_NUM;
        }
        setSheet(getWorkbook().createSheet(
                sheetName.replace(ExcelWriter.SHEET_NUM,
                        String.valueOf((++sheetIndex) + 1))
        ));

        // 重置列索引
        setRowIndex(-1);
        // 重置行索引
        super.setCellIndex(-1);
        if (sheetIndex > 0) {
            // 渲染之前的工作簿
            lazyRender(getSheet());
        }
        setSheetLazyRenderFlag(false);
    }

    public String getDefaultSheetName() {
        return defaultSheetName;
    }

    public void setDefaultSheetName(String defaultSheetName) {
        this.defaultSheetName = defaultSheetName;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }
}
