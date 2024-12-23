package com.example.easyexcel_showcase.handler;

import com.alibaba.excel.metadata.data.CellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

@Slf4j
public class CustomCellWriteHandler implements CellWriteHandler {
    private Integer dataSize;
    private Integer titleSize;
    private Integer insuranceTitleSize;

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        int rowIndex = context.getCell().getRowIndex();
        int columnIndex = context.getCell().getColumnIndex();

        WriteCellData<?> cellData = context.getFirstCellData();
        // 这里需要去cellData 获取样式
        WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();
        //设置样式

        //居中对齐
        writeCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //字体
        WriteFont font = new WriteFont();
        writeCellStyle.setWriteFont(font);
        font.setFontName("宋体");

        //获取cell内容(获取不到head内容)
        Object originalValue = context.getOriginalValue();
        System.out.printf("rowIndex:%d, columnIndex:%d value: %s%n", rowIndex, columnIndex, originalValue != null ? originalValue.toString():"null");

        //head逻辑
        if(context.getHead()) {
            //标题样式设置
            if(rowIndex == 0) {
                writeCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
                font.setFontHeightInPoints((short)21);
                //设置行高
                context.getRow().setHeightInPoints((short)25);
            }else{
                writeCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
            }
        }else{
            //合计的合并
            if(dataSize + 1 == rowIndex && columnIndex == 0){
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
                context.getCell().getSheet().addMergedRegion(cellRangeAddress);
            }

            //最后一行的合并(不同sheet的列数不一样, 需要分情况处理)
            //制表：                 审核：                      部门负责人：
            Integer sheetNo = context.getWriteSheetHolder().getSheetNo();
            if(sheetNo == 0 && dataSize + 2 == rowIndex && columnIndex == 0){
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, titleSize-1);
                context.getCell().getSheet().addMergedRegion(cellRangeAddress);
                writeCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
            }else if(sheetNo == 1 && dataSize + 2 == rowIndex && columnIndex == 0) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, insuranceTitleSize - 1);
                context.getCell().getSheet().addMergedRegion(cellRangeAddress);
                writeCellStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
            }
        }
    }

    public Integer getDataSize() {
        return dataSize;
    }

    public void setDataSize(Integer dataSize) {
        this.dataSize = dataSize;
    }

    public Integer getTitleSize() {
        return titleSize;
    }

    public void setTitleSize(Integer titleSize) {
        this.titleSize = titleSize;
    }

    public Integer getInsuranceTitleSize() {
        return insuranceTitleSize;
    }

    public void setInsuranceTitleSize(Integer insuranceTitleSize) {
        this.insuranceTitleSize = insuranceTitleSize;
    }
}