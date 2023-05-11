package com.wz.utils;


import com.wz.annotation.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @author 王哲
 * @ClassName ExcelUtils
 * @create 2023--五月--下午9:13
 * @Description ExcelUtil工具类，用于实现通用的Excel导入和导出功能。
 * @Version V1.0
 */
public class ExcelUtils {
    /**
     * 导入Excel文件数据到对象列表中。
     *
     * @param file 上传的文件
     * @param clazz 对象类型
     * @param <T>   对象泛型
     * @return 对象列表
     * @throws Exception 异常
     */
    public static <T> List<T> importExcel(MultipartFile file, Class<T> clazz) throws Exception {
        List<T> list = new ArrayList<>();
        Workbook workbook = getWorkbook(file);
        Sheet sheet = workbook.getSheetAt(0);
        Field[] fields = clazz.getDeclaredFields();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            T obj = clazz.getDeclaredConstructor().newInstance();
            Row row = sheet.getRow(i);

            for (int j = 0; j < fields.length; j++) {
                String value = getStringCellValue(row.getCell(j));
                setFieldValue(obj, fields[j], value);
            }
            list.add(obj);
        }
        return list;
    }

    /**
     * 导出对象列表数据到Excel文件中。
     *
     * @param response HTTP响应对象
     * @param list     对象列表
     * @param clazz    对象类型
     * @param <T>      对象泛型
     * @throws Exception 异常
     */
    public static <T> void exportExcel(HttpServletResponse response, List<T> list, Class<T> clazz) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        sheet.setDefaultColumnWidth(20);
        Row titleRow = sheet.createRow(0);
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelExportField> exportFields = getExportFields(fields);

        for (int i = 0; i < exportFields.size(); i++) {
            titleRow.createCell(i).setCellValue(exportFields.get(i).getTitle());
        }

        for (int i = 0; i < list.size(); i++) {
            Row row = sheet.createRow(i + 1);
            T obj = list.get(i);

            for (int j = 0; j < exportFields.size(); j++) {
                Object value = getFieldValue(obj, fields, exportFields.get(j).getName());
                row.createCell(j).setCellValue(value.toString());
            }
        }

        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment;filename=data.xlsx");
        workbook.write(response.getOutputStream());
    }
    /**
     * 获取Excel工作簿对象。
     *
     * @param file 上传的文件
     * @return Excel工作簿对象
     * @throws Exception 异常
     */
    private static Workbook getWorkbook(MultipartFile file) throws Exception {
        String fileName = file.getOriginalFilename();
        if (fileName.endsWith(".xls")) {
            return new HSSFWorkbook(file.getInputStream());
        } else if (fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(file.getInputStream());
        }
        throw new Exception("Unsupported file type.");
    }
    /**
     * 获取单元格的String类型数据。
     *
     * @param cell 单元格对象
     * @return 单元格String类型数据
     */
    private static String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        cell.setCellType(CellType.STRING);
        return cell.getStringCellValue();
    }
/**
 * 设置对象的属性值。
 * *
 * * @param obj 对象实例
 * * @param field 属性对象
 * * @param value 属性值
 * * @param <T> 对象泛型
 * * @throws Exception 异常
 * */
    private static <T> void setFieldValue(T obj, Field field, String value) throws Exception {
        field.setAccessible(true);
        if (field.getType() == String.class) {
            field.set(obj, value);
        } else if (field.getType() == Integer.class || field.getType() == int.class) {
            field.set(obj, Integer.valueOf(value));
        } else if (field.getType() == Long.class || field.getType() == long.class) {
            field.set(obj, Long.valueOf(value));
        } else if (field.getType() == Float.class || field.getType() == float.class) {
            field.set(obj, Float.valueOf(value));
        } else if (field.getType() == Double.class || field.getType() == double.class) {
            field.set(obj, Double.valueOf(value));
        } else if (field.getType() == BigDecimal.class) {
            field.set(obj, new BigDecimal(value));
        } else {
            throw new Exception("Unsupported data type");
        }
    }
    /**
     * 获取对象的属性值。
     *
     * @param obj     对象实例
     * @param fields  类属性数组
     * @param name    属性名字
     * @param <T>     对象泛型
     * @return 对象属性值
     * @throws Exception 异常
     */
    private static <T> Object getFieldValue(T obj, Field[] fields, String fieldName) throws Exception {
        for (Field field : fields) {
            if (field.getName().equals(fieldName)) {
                field.setAccessible(true);
                return field.get(obj);
            }
        }
        throw new Exception("Field not found.");
    }
    /**
     * 获取需要导出的属性列表。
     *
     * @param fields 类属性数组
     * @return 需要导出的属性列表
     */
    private static List<ExcelExportField> getExportFields(Field[] fields) {
        List<ExcelExportField> exportFields = new ArrayList<>();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                exportFields.add(new ExcelExportField(excelField.title(), field.getName(), excelField.order()));
            }
        }
        Collections.sort(exportFields);
        return exportFields;
    }
    /**
     * Excel导出属性对象，用于保存标注了ExcelField注解的属性的信息。
     */
    static class ExcelExportField implements Comparable<ExcelExportField> {
        private String title;
        private String name;
        private int order;

        public ExcelExportField(String title, String name, int order) {
            this.title = title;
            this.name = name;
            this.order = order;
        }public String getTitle() {
            return title;
        }

        public String getName() {
            return name;
        }

        public int getOrder() {
            return order;
        }

        @Override
        public int compareTo(ExcelExportField o) {
            return Integer.compare(this.order, o.getOrder());
        }
    }
}
