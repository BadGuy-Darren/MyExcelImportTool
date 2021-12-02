package com.darren.tools.getexcelutil;

import com.darren.tools.getexcelutil.annotation.ValueIngnore;
import com.darren.tools..getexcelutil.annotation.NotNull;
import com.darren.tools.getexcelutil.annotation.ExcelDateFormat;
import com.darren.tools.getexcelutil.annotation.ExcelNumberFormat;
import com.darren.tools.getexcelutil.annotation.ValueLimit;
import com.darren.tools.getexcelutil.exceptions.EmptyExcelFileException;
import com.darren.tools.getexcelutil.exceptions.IllegalStatementsException;
import com.darren.tools.getexcelutil.exceptions.SheetNoOuntOfBoundsException;
import com.darren.tools.getexcelutil.exceptions.SheetNumOuntOfBoundsException;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.ArrayList;
import java.util.Set;
import java.util.HashSet;
import java.util.Arrays;
import java.util.Date;

/**
 * @code 大迪
 * date 2021/09/22
 * @version 4.1
 * Description - GetExcelUtil 4.1 - 瘦身版
 * IllegalArgumentException（非法参数）用于抛出由于用户操作不合规抛出的异常
 */
public class GetExcelUtil4 implements Serializable {
    private static final long serialVersionUID = 1L;
    private static final Logger logger = LoggerFactory.getLogger(GetExcelUtil4.class);

    /**
     * 參數
     */
    private String datePattern = "yyyy/MM/dd";
    private String numberFormat = "#.##";
    private Sheet[] sheets;
    private String startTag;
    private String endTag;

    /**
     * 初始化sheet
     * @param file
     *  - 要读取数据的excel文件
     * @param sheetNum
     *  - 欲读取的工作表数
     * @return this
     * @throws IOException
     *  - 开启文件输入流异常
     */
    public GetExcelUtil4 initialize(MultipartFile file, int sheetNum) throws IOException {
        if (null == file) {
            throw new EmptyExcelFileException("读取的文件不存在！");
        }
        String fileName = file.getOriginalFilename();
        assert fileName != null;
        ZipSecureFile.setMinInflateRatio(-1.0d);
        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            throw new IllegalArgumentException("未预期的文件格式：" + fileName);
        }

        // Mark try-with-resource机制，资源关闭交由java管理 （资源类需实现Closeable或AutoCloseable接口）
        try (BufferedInputStream inputStream = new BufferedInputStream(file.getInputStream());
             Workbook workbook = createWorkbook(inputStream)) {
            if (sheetNum > workbook.getNumberOfSheets()) {
                throw new SheetNumOuntOfBoundsException("文件工作表个数小于欲读取的数量！");
            }
            sheets = new Sheet[sheetNum];
            for (int i = 0; i < sheetNum; i++) {
                // 有隐藏的sheet不读取
                if (workbook.getSheetVisibility(i).equals(SheetVisibility.HIDDEN)) {
                    throw new IllegalArgumentException("不允许存在隐藏的工作表！");
                }
                sheets[i] = workbook.getSheetAt(i);
            }
        }
        return this;
    }

    private Workbook createWorkbook(BufferedInputStream bis) throws IOException {
        switch (FileMagic.valueOf(bis)) {
            case OOXML:
                return new XSSFWorkbook(bis);
            case OLE2:
                return new HSSFWorkbook(bis);
            default:
                throw new IllegalArgumentException("未预期的文件格式：" + FileMagic.valueOf(bis));
        }
    }


    /**
     * 设置获取日期格式
     * @param datePattern
     *  - 默认yyyy/MM/dd
     * @return this
     */
    public GetExcelUtil4 setDatePattern(String datePattern) {
        this.datePattern = datePattern;
        return this;
    }


    /**
     * 设置获取的数字格式
     * @param numberFormat
     *  - 默认#.####
     * @return this
     */
    public GetExcelUtil4 setNumberFormat(String numberFormat) {
        this.numberFormat = numberFormat;
        return this;
    }


    /**
     * 设置开始标志，采用批注
     * @param startTag
     * @return this
     */
    public GetExcelUtil4 setStartTag(String startTag) {
        this.startTag = startTag;
        return this;
    }


    /**
     * 设置结束标志
     * @param endTag
     * @return this
     */
    public GetExcelUtil4 setEndTag(String endTag) {
        this.endTag = endTag;
        return this;
    }


    /**
     * 获取sheet名
     * @param sheetNo 第几个sheet (1-based)
     * @return sheet名
     */
    public String getSheetName(int sheetNo) {
        if (sheetNo > sheets.length) {
            throw new SheetNoOuntOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }
        return sheets[sheetNo - 1].getSheetName();
    }

    /**
     * 获取数据水平分布的表格的内容
     * @param sheetNo 第几个sheet
     * @param startRow 起始行
     * @param startColumn 起始列
     * @param cellNum 一行数据量
     * @param clazz 接收实体类型
     * @param <T> 实体泛型
     * @throws IllegalStatementsException
     *      {@link com.darren.tools.getexcelutil.annotation.ExcelNumberFormat}注解的when条件判定时，表达式格式不对，
     *      或表达式字段不存在时会抛出的异常
     * @throws SheetNoOuntOfBoundsException
     *      读取的sheet序号溢出
     * @throws EmptyExcelFileException
     *
     *
     * setField方法抛出的
     * @return 目标sheet数据
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> getHorizontalData(int sheetNo, int startRow, int startColumn, int cellNum, Class<T> clazz) throws IllegalStatementsException {
        if (sheetNo > sheets.length) {
            throw new SheetNoOuntOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }

        if (StringUtils.isNotEmpty(startTag) && !startTag.equals(getComment(sheetNo, startRow, startColumn))) {
            throw new IllegalArgumentException("无法适配起始行！");
        }

        List<T> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        if (isBaseType(clazz)) {
            for (int rowNum = startRow; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                Row row = sheet.getRow(rowNum - 1);
                Cell cell = row.getCell(startColumn - 1);
                String cellVal = getCellVal(cell);
                list.add((T) ConvertUtils.convert(cellVal, clazz));
            }
        } else {
            for (int rowNum = startRow; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                Row row = sheet.getRow(rowNum - 1);

                // 有结束标志，且行第一个值为结束标志时结束
                if (StringUtils.isNotEmpty(endTag) && endTag.equals(getCellVal(row.getCell(startColumn - 1)))) {
                    break;
                }

                T t;
                try {
                    t = clazz.newInstance();
                } catch (InstantiationException | IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
                Field[] fields = clazz.getDeclaredFields();
                int fieldsLen = fields.length;
                int fieldIndex = 0;
                int columnNum = startColumn;

                // 忽略接收实体类中的字段
                while (fieldIndex < fieldsLen) {
                    if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                        fieldIndex++;
                        continue;
                    }
                    if (columnNum < startColumn + cellNum) {
                        // 字段赋值
                        setField(t, clazz, fields[fieldIndex], row.getCell(columnNum - 1), sheetNo, rowNum, columnNum);
                        columnNum++;
                        fieldIndex++;
                    }
                }

                // 全空行时跳出
                if (!ObjectUtils.anyNotNull(t)) {
                    break;
                }
                list.add(t);
            }
        }

        if (list.size() == 0) {
            throw new EmptyExcelFileException("文件读取数据区域内容为空！");
        }

        return list;
    }

    /**
     * 获取数据垂直分布的表格的内容
     * @param sheetNo 第几个sheet
     * @param startRow 起始行
     * @param startColumn 起始列
     * @param cellNum 一列数据量
     * @param clazz 接收实体类型
     * @param <T> 实体泛型
     * @return 目标sheet数据
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> getVerticalData(int sheetNo, int startRow, int startColumn, int cellNum, Class<T> clazz) throws IllegalStatementsException {
        if (sheetNo > sheets.length) {
            throw new SheetNoOuntOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }

        if (StringUtils.isNotEmpty(startTag) && startTag.equals(getComment(sheetNo, startRow, startColumn))) {
            throw new IllegalArgumentException("无法适配起始行！");
        }

        List<T> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        if (isBaseType(clazz)) {
            for (int columnNum = startColumn; columnNum <= sheet.getRow(startRow - 1).getLastCellNum(); columnNum++) {
                Row row = sheet.getRow(startRow - 1);
                Cell cell = row.getCell(columnNum - 1);
                String cellVal = getCellVal(cell);

                list.add((T) ConvertUtils.convert(cellVal, clazz));
            }
        } else {
            // 遍历sheet中的列
            for (int columnNum = startColumn; columnNum <= sheet.getRow(startRow - 1).getLastCellNum(); columnNum++) {
                // 有结束标志，且行第一个值为结束标志时结束
                if (StringUtils.isNotEmpty(endTag) && endTag.equals(getCellVal(sheet.getRow(startRow - 1).getCell(columnNum - 1)))) {
                    break;
                }

                T t;
                try {
                    t = clazz.newInstance();
                } catch (InstantiationException | IllegalAccessException e) {
                    throw new RuntimeException(e);
                }
                Field[] fields = clazz.getDeclaredFields();
                int fieldsLen = fields.length;
                int fieldIndex = 0;
                int rowNum = startRow;

                while (fieldIndex < fieldsLen) {
                    // 忽略接收实体类中的字段
                    if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                        fieldIndex++;
                        continue;
                    }
                    if (rowNum < startRow + cellNum) {
                        setField(t, clazz, fields[fieldIndex], sheet.getRow(rowNum - 1).getCell(columnNum - 1), sheetNo, rowNum, columnNum);
                        rowNum++;
                        fieldIndex++;
                    }
                }

                // 全空行时跳出
                if (!ObjectUtils.anyNotNull(t)) {
                    break;
                }
                list.add(t);
            }
        }

        if (list.size() == 0) {
            throw new EmptyExcelFileException("文件读取数据区域内容为空！");
        }

        return list;
    }


    /**
     * 获取批注
     * @param rowNo
     *  - 行号
     * @param columnNo
     *  - 列号
     * @return 对应位置的批注
     */
    private String getComment(int sheetNo, int rowNo, int columnNo) {
        if (sheetNo > sheets.length) {
            throw new SheetNoOuntOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }
        return sheets[sheetNo - 1].getRow(rowNo - 1).getCell(columnNo - 1).getCellComment().getString().toString();
    }


    /**
     * 实体类赋值
     *  包含注解防呆
     * @param o - 进行变量赋值的实体类
     * @param classType - 实体类的Class
     * @param field - 赋值字段
     * @param cell - 对应的Excel单元
     * @param rowNum - 行序号
     * @param columnNum - 列序号
     */
    private void setField(Object o, Class<?> classType, Field field, Cell cell, int sheetNo, int rowNum, int columnNum) throws IllegalStatementsException {
        String cellVal = getCellVal(cell, field);
        if (field.isAnnotationPresent(NotNull.class)) {
            NotNull valuePattern = field.getAnnotation(NotNull.class);
            if (StringUtils.isEmpty(cellVal)) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnNum + "列：" + valuePattern.message());
            }
        }
        if (field.isAnnotationPresent(ExcelDateFormat.class)) {
            ExcelDateFormat edf = field.getAnnotation(ExcelDateFormat.class);
            SimpleDateFormat sdf = new SimpleDateFormat(edf.pattern());
            try {
                sdf.parse(cellVal);
            } catch (ParseException e) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnNum + "列：" + edf.message());
            }
        }
        if (field.isAnnotationPresent(ExcelNumberFormat.class)) {
            ExcelNumberFormat enf = field.getAnnotation(ExcelNumberFormat.class);
            String when = enf.when();
            if (StringUtils.isEmpty(when) || isWhen(o, classType, when)) {
                DecimalFormat df = new DecimalFormat(enf.format());
                try {
                    df.format(new BigDecimal(cellVal));
                } catch (Exception e) {
                    logger.error("", e);
                    throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnNum + "列：" + enf.message());
                }
            }
        }
        if (field.isAnnotationPresent(ValueLimit.class)) {
            ValueLimit vl = field.getAnnotation(ValueLimit.class);
            String[] limit = vl.limit();
            Set<String> set = new HashSet<>(Arrays.asList(limit));
            if (!set.contains(cellVal)) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnNum + "列：" + vl.message());
            }
        }
        // 赋值
        setFieldValue(o, classType, field, cellVal);
    }

    private boolean isWhen(Object target, Class<?> classType, String when) throws IllegalStatementsException {
        if (!when.contains("==")) {
            throw new IllegalStatementsException("com.darren.tools.getexcelutil.annotation.ExcelNumberFormat.when()：缺失条件连接符“==”");
        }

        String caseFieldName = when.substring(0, when.indexOf("==")).trim();
        String caseValue = when.substring(when.indexOf("==") + 2).trim();

        if (StringUtils.isEmpty(caseFieldName)) {
            throw new IllegalStatementsException("com.darren.tools.getexcelutil.annotation.ExcelNumberFormat.when()：条件判定字段不能为空！");
        }
        if (StringUtils.isEmpty(caseValue)) {
            throw new IllegalStatementsException("com.darren.tools.getexcelutil.annotation.ExcelNumberFormat.when()：条件判定值不能为空！");
        }

        String getMethodName = "get" + caseFieldName.substring(0, 1).toUpperCase() + caseFieldName.substring(1);
        Method getMethod;
        try {
            getMethod = classType.getDeclaredMethod(getMethodName);
        } catch (NoSuchMethodException e) {
            throw new IllegalStatementsException("未知字段：" + caseFieldName);
        }
        String value;
        // 赋值
        try {
            value = (String) getMethod.invoke(target);
        } catch (IllegalAccessException | InvocationTargetException e) {
            logger.error("", e);
            value = "";
        }

        return caseValue.equals(value);
    }

    private void setFieldValue(Object target, Class<?> classType, Field field, Object value) {
        // 获取参数类
        Class<?>[] paramTypes = new Class[1];
        paramTypes[0] = field.getType();

        // 转化值
        Object[] values = new Object[1];
        values[0] = ConvertUtils.convert(value, paramTypes[0]);

        // 获取setter
        String fieldName = field.getName();
        String setNameMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
        Method setMethod;
        try {
            setMethod = classType.getDeclaredMethod(setNameMethodName, paramTypes);
        } catch (NoSuchMethodException e) {
            throw new RuntimeException(e);
        }

        // 赋值
        try {
            setMethod.invoke(target, values);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    private String getCellVal(Cell cell) {
        return getCellVal(cell, null);
    }

    private String getCellVal(Cell cell, Field field) {
        if (cell == null) {
            return "";
        }

        String cellString;

        switch (cell.getCellType()) {
            case STRING: // 字符串
                cellString = cell.getStringCellValue();
                break;
            case NUMERIC: // 数字
            case FORMULA:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //用于转化为日期格式
                    String pattern = (null != field && field.isAnnotationPresent(ExcelDateFormat.class)) ?
                            field.getAnnotation(ExcelDateFormat.class).pattern() : datePattern;
                    Date d = cell.getDateCellValue();
                    DateFormat f = new SimpleDateFormat(pattern);
                    cellString = f.format(d);
                } else {
                    // 用于格式化数字，只保留两位小数
                    String format = (null != field && field.isAnnotationPresent(ExcelDateFormat.class)) ?
                            field.getAnnotation(ExcelNumberFormat.class).format() : numberFormat;
                    DecimalFormat df = new DecimalFormat(format);
                    cellString = df.format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN: // Boolean
                cellString = String.valueOf(cell.getBooleanCellValue());
                break;
//            case FORMULA: // 公式
//                cellString = String.valueOf(cell.getNumericCellValue());
//                // cellString = cell.getStringCellValue();
//                break;
            case BLANK: // 空值
            case ERROR: // 故障
                cellString = "";
                break;
            default:
                cellString = "ERROR";
                break;
        }
        return cellString.trim();
    }

    private boolean isBaseType (Class<?> clazz) {
        return String.class.equals(clazz)
                || Integer.class.equals(clazz)
                || Double.class.equals(clazz)
                || Object.class.equals(clazz)
                || Long.class.equals(clazz);
    }

}
