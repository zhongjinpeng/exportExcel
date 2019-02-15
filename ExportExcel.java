package com.ld.admin.survey.utils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;
import java.util.stream.IntStream;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

import com.ld.utils.Identities;
import com.ld.utils.ObjectUtils;

/**
 * 
 * @author zhongjinpeng
 * @date 2018年09月03日
 * @Description: 导出excel工具类
 */
public interface ExportExcel<T> {
    
    /**
     * 将字节数组写出到servlet输出流
     * 
     * @param response http回应对象，为excel回应的目的地
     * @param list list 要导出到 excel的数据集合
     * @param titles excel的标题 通常取第一行作为excel的标题
     * @param columnWidth 列宽参数 若为空则用默认值
     * @param fileName 文件名 若未空则以uuid作为文件名
     * @param b 图片数据
     * @throws IOException
     */
    default void exportExcel(HttpServletResponse response, List<T> list, String[] titles, int[] columnWidth,
            String fileName,byte[] b) throws IOException {
        if (ObjectUtils.isEmpty(fileName)) {
            fileName = Identities.uuid2();
        }
        byte[] bytes = selectExcel(list, titles, columnWidth, fileName, b);
        
//        response.setContentType("APPLICATION/vnd.ms-excel;charset=UTF-8");
        response.setContentType("application/x-msdownload");
        response.setCharacterEncoding("UTF-8");
        response.setHeader("Content-Disposition",
                "attachment; filename=".concat(String.valueOf(URLEncoder.encode(fileName + ".xlsx", "UTF-8"))));
        response.setContentLength(bytes.length);
        response.getOutputStream().write(bytes);
        response.getOutputStream().flush();
        response.getOutputStream().close();
    }

    /**
     * 选择要导出的文件 导出的excel 属于office 2007格式的文件
     * 
     * @param list excel文件内容
     * @param titles excel 文件的标题
     * @param columnWidth 列宽
     * @param fileName 文件名
     * @return 已经生成excel文件的字节数组
     * @throws IOException
     */
    default byte[] selectExcel(List<T> list, String[] titles, int[] columnWidth, String fileName,byte[] b) throws IOException {
        Workbook workbook = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet();
         
        generateExcelTitle(titles, sheet, workbook, fileName);
        eachListAndCreateRow(list, sheet, titles, columnWidth);
        if(ObjectUtils.isNotEmpty(b)) {
        	 insertChart(workbook,sheet,b);
        }
//        Sheet sheet2 = workbook.createSheet();
//        
//        generateExcelTitle(titles, sheet2, workbook, fileName);
//        eachListAndCreateRow(list, sheet2, titles, columnWidth);
//        if(ObjectUtils.isNotEmpty(b)) {
//             insertChart(workbook,sheet2,b);
//        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        return out.toByteArray();
    }
    
    /**
     * 插入图表
     * @param workbook
     * @param sheet
     * @param bos
     * @throws IOException 
     */
    default void insertChart (Workbook workbook,Sheet sheet,byte[] b) throws IOException {
  	    // 画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
		Drawing<?> patriarch = sheet.createDrawingPatriarch();
		// anchor主要用于设置图片的属性(1,1 表示图片左上角，15,31表示图片右下角，通过这个可以控制图片的大小)
		XSSFClientAnchor anchor = new XSSFClientAnchor(0, 150, 1000, 210, (short) 10, 10, (short) 24, 40);
		anchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE);
		// 插入图片
		patriarch.createPicture(anchor, workbook.addPicture(b, HSSFWorkbook.PICTURE_TYPE_JPEG));
    }
    
    /**
     * 遍历集合，并创建单元格行
     * 
     * @param list 数据集合
     * @param sheet 工作簿
     * @param titles excel 文件的标题
     * @param columnWidth 列宽
     */
    default void eachListAndCreateRow(List<T> list, Sheet sheet, String[] titles, int[] columnWidth) {
        // 设置列宽
        if (ObjectUtils.isNotEmpty(columnWidth)) {
            for (int i = 0; i < titles.length; i++) {
                for (int j = 0; j <= i; j++) {
                    if (i == j) {
                        sheet.setColumnWidth(i, columnWidth[j] * 256); // 单独设置每列的宽
                    }
                }
            }
        }
        IntStream.range(0, list.size()).forEach(i -> generateExcelForAs(list.get(i), sheet.createRow((i + 2))));
    }

    /**
     * 生成excel文件的标题以及表头
     * 
     * @param titles excel 文件的标题
     * @param sheet 工作簿
     * @param workbook
     * @param fileName 文件名
     */
    default void generateExcelTitle(String[] titles, Sheet sheet, Workbook workbook, String fileName) {

        // 创建标题
        Row headline = sheet.createRow(0);
        Cell title = headline.createCell(0);// 创建标题第一列
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, titles.length - 1)); // 合并列标题
        title.setCellValue(fileName); // 设置值标题
        title.setCellStyle(setHeadLineStyle(workbook)); // 设置标题样式

        // 创建第1行 也就是表头
        Row header = sheet.createRow((int) 1);
        header.setHeightInPoints(37);// 设置表头高度
        for (int i = 0; i < titles.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(setContentStyle(workbook));
        }
    }
    
    /**
     * 设置内容和表头的样式
     * @param workbook
     * @return
     */
    default CellStyle setContentStyle(Workbook workbook) {
        CellStyle contenStyle = workbook.createCellStyle();
        contenStyle.setWrapText(true);// 设置自动换行
        contenStyle.setAlignment(HorizontalAlignment.CENTER);
        contenStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 创建一个居中格式

//        contenStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
//        contenStyle.setBorderBottom(BorderStyle.THIN);
//        contenStyle.setBorderLeft(BorderStyle.THIN);
//        contenStyle.setBorderRight(BorderStyle.THIN);
//        contenStyle.setBorderTop(BorderStyle.THIN);

        Font headerFont = (Font) workbook.createFont(); // 创建字体样式
        headerFont.setBold(true); // 字体加粗
        headerFont.setFontName("黑体"); // 设置字体类型
        headerFont.setFontHeightInPoints((short) 10); // 设置字体大小
        contenStyle.setFont(headerFont); // 为标题样式设置字体样式
        return contenStyle;
    }
    
    /**
     * 设置标题的样式
     * @param workbook
     * @return
     */
    default CellStyle setHeadLineStyle(Workbook workbook) {
        CellStyle headLineStyle = workbook.createCellStyle();
        headLineStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());// 设置单元格着色
        headLineStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND); // 设置单元格填充样式
        headLineStyle.setBorderBottom(BorderStyle.THIN);// 设置下边框
        headLineStyle.setBorderLeft(BorderStyle.THIN);// 设置左边框
        headLineStyle.setBorderRight(BorderStyle.THIN);// 设置右边框
        headLineStyle.setBorderTop(BorderStyle.THIN);// 上边框
        headLineStyle.setAlignment(HorizontalAlignment.CENTER);// 居中
        // 生成字体
        Font headLineFont = workbook.createFont(); // 创建字体样式
        headLineFont.setColor(IndexedColors.WHITE.getIndex());
        headLineFont.setFontHeightInPoints((short) 15);
        headLineFont.setBold(true);
        headLineFont.setFontName("黑体");
        headLineStyle.setFont(headLineFont); // 为标题样式设置字体样式
        return headLineStyle;
    }
    
    /**
     * 创建excel内容文件
     * 
     * @param t 组装excel 文件的内容
     * @param row 当前excel 工作行
     */
    void generateExcelForAs(T t, Row row);

}
