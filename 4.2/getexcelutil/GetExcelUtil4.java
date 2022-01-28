package com.foxconn.indint.utils.getexcelutil;

import com.foxconn.indint.utils.getexcelutil.annotation.DynamicRank;
import com.foxconn.indint.utils.getexcelutil.annotation.Transform;
import com.foxconn.indint.utils.getexcelutil.annotation.ValueIngnore;
import com.foxconn.indint.utils.getexcelutil.annotation.NotNull;
import com.foxconn.indint.utils.getexcelutil.annotation.ExcelDateFormat;
import com.foxconn.indint.utils.getexcelutil.annotation.ExcelNumberFormat;
import com.foxconn.indint.utils.getexcelutil.annotation.ValueLimit;
import com.foxconn.indint.utils.getexcelutil.exceptions.DataDuplicationException;
import com.foxconn.indint.utils.getexcelutil.exceptions.EmptyExcelFileException;
import com.foxconn.indint.utils.getexcelutil.exceptions.IllegalStatementsException;
import com.foxconn.indint.utils.getexcelutil.exceptions.SheetNoOutOfBoundsException;
import com.foxconn.indint.utils.getexcelutil.exceptions.SheetNumOutOfBoundsException;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
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
 * date 2022/01/18
 * @version 4.2
 * Description - GetExcelUtil 4.2
 * 1. 增加@DynamicRank注解功能，支持获取动态表头及对应栏位。
 *    目前仅支持动态表头在固定表头之后的情况
 * 2. 增加支持重复性检测
 * 3. 增加@Transform注解功能，支持字段值的转换
 * 4. 优化：
 *      （1）注解数值格式(@ExcelNumberFormat)时，字段遇空值自动转化0
 *      （2）设置起始标志的情况下，校准位置改为以标题栏为准（取数据行前推一行）
 *      （3）异常信息提示的数字列序号校准为对应字母序号
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
                throw new SheetNumOutOfBoundsException("文件工作表个数小于欲读取的数量！");
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
            throw new SheetNoOutOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }
        return sheets[sheetNo - 1].getSheetName();
    }


    /**
     * 数据重复性检查
     * desc:
     *  1. 依赖于实体类equals方法的实现
     *  2. 需要遍历已获取到的数据集，会影响取值效率，非必要不开启
     *
     * @param checkRepeat
     */

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
            throw new SheetNoOutOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }
        CellAddress cellAddress = new CellAddress(rowNo - 1, columnNo - 1);
        return sheets[sheetNo - 1].getCellComment(cellAddress).getString().toString();
    }


    /**
     * 获取数据水平分布的表格的内容
     * 默认不进行重复性检查
     */
    public <T> List<T> getHorizontalData(int sheetNo, int startRow, int startColumn, Class<T> clazz) throws IllegalStatementsException {
        return getHorizontalData(sheetNo, startRow, startColumn, clazz, false);
    }


    /**
     * 获取数据水平分布的表格的内容
     * @param sheetNo 第几个sheet
     * @param startRow 起始行
     * @param startColumn 起始列
     * @param clazz 接收实体类型
     * @param enableDuplicateCheck 数据重复性检查
     *        1. 依赖于实体类equals方法的实现
     *        2. 需要遍历已获取到的数据集，会影响取值效率，非必要不开启
     * @param <T> 实体泛型
     * @throws IllegalStatementsException
     *      {@link com.foxconn.indint.utils.getexcelutil.annotation.ExcelNumberFormat}注解的when条件判定时，表达式格式不对，
     *      或表达式字段不存在时会抛出的异常
     * @throws SheetNoOutOfBoundsException
     *      读取的sheet序号溢出
     * @throws EmptyExcelFileException
     *
     *
     * setField方法抛出的
     * @return 目标sheet数据
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> getHorizontalData(int sheetNo, int startRow, int startColumn, Class<T> clazz, boolean enableDuplicateCheck) throws IllegalStatementsException {
        if (sheetNo > sheets.length) {
            throw new SheetNoOutOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }

        String startCommon;
        try {
            startCommon = getComment(sheetNo, startRow - 1, startColumn);
        } catch (NullPointerException e) {
            throw new IllegalArgumentException("无法适配起始行！");
        }
        if (StringUtils.isNotEmpty(startTag) && !startTag.equals(startCommon)) {
            throw new IllegalArgumentException("无法适配起始行！");
        }

        List<T> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        if (isBaseType(clazz)) {
            for (int rowNum = startRow; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                Row row = sheet.getRow(rowNum - 1);
                Cell cell = row.getCell(startColumn - 1);

                if (isEmptyCell(cell)) break;

                String cellVal = getCellVal(cell);

                // 重复性检查
                if (enableDuplicateCheck && list.size() > 0) {
                    for (int i = 0; i < list.size(); i ++) {
                        if (cellVal.equals(list.get(i).toString())) {
                            throw new DataDuplicationException("文件工作表" + sheetNo + "第" + (i + startRow) + "行与第" + rowNum + "行重复！");
                        }
                    }
                }

                list.add((T) ConvertUtils.convert(cellVal, clazz));
            }
        } else {
            for (int rowNum = startRow; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
                Row row = sheet.getRow(rowNum - 1);
                Field[] fields = clazz.getDeclaredFields();
                int fieldsLen = fields.length;

                // 空行退出
                if (isEmptyRow(row)) break;

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
                int fieldIndex = 0;
                int columnNum = startColumn;

                while (fieldIndex < fieldsLen) {
                    // 忽略接收实体类中的字段，跳过数组长度辅助字段
                    if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                        fieldIndex++;
                        continue;
                    }
                    // 字段赋值
                    int assignedCellNum = setHorizontalField(t, clazz, fields[fieldIndex], row.getCell(columnNum - 1), sheetNo, rowNum, columnNum);
                    columnNum += assignedCellNum;
                    fieldIndex ++;
                }

                // 重复性检查
                if (enableDuplicateCheck && list.size() > 0) {
                    for (int i = 0; i < list.size(); i ++) {
                        if (t.equals(list.get(i))) {
                            throw new DataDuplicationException("文件工作表" + sheetNo + "第" + (i + startRow) + "行与第" + rowNum + "行重复！");
                        }
                    }
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
     * 实体类赋值
     *  包含注解防呆
     * @param o - 进行变量赋值的实体类
     * @param classType - 实体类的Class
     * @param field - 赋值字段
     * @param cell - 对应的Excel单元
     * @param rowNum - 行序号
     * @param columnNum - 列序号
     * @return 操作cell数
     */
    private int setHorizontalField(Object o, Class<?> classType, Field field, Cell cell, int sheetNo, int rowNum, int columnNum) throws IllegalStatementsException {
        Object fieldValue;
        int assignedCell;

        if (field.isAnnotationPresent(DynamicRank.class)) {
            ParameterizedType types = (ParameterizedType) field.getGenericType();
            Type type = types.getActualTypeArguments()[0];

            DynamicRank dynamicRank = field.getAnnotation(DynamicRank.class);
            int[] rows = {dynamicRank.titleRank(), rowNum};
            boolean enableDuplicateCheck = dynamicRank.enableDuplicateCheck();

            List<Object> list = getHorizontalDynamicRank(sheetNo, rows, columnNum, (Class<?>) type, enableDuplicateCheck);
            fieldValue = list;
            assignedCell = list.size();
        } else {
            // 获取通过注解校验的值
            fieldValue = getVerifiedCellVal(o, classType, field, cell, sheetNo, rowNum, columnNum);
            assignedCell = 1;
        }

        // 赋值
        setFieldValue(o, classType, field, fieldValue);

        return assignedCell;
    }


    /**
     * 获取水平表格动态栏位
     * @param sheetNo sheet序号
     * @param rows 要获取动态栏位的行（必包括当前行，可包括表头栏等）
     * @param startColumn 开始列
     * @param clazz 实体类类型
     * @param enableDuplicateCheck 重复性检查，依赖于实体类equals方法的实现
     * @return 动态栏位数据
     * @throws IllegalStatementsException 必解决异常
     */
    private List<Object> getHorizontalDynamicRank(int sheetNo, int[] rows, int startColumn, Class<?> clazz, boolean enableDuplicateCheck) throws IllegalStatementsException {
        List<Object> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        // -1表示不限制，取到底
        int endColumn = sheet.getRow(rows[0] - 1).getLastCellNum();

        Field[] fields = clazz.getDeclaredFields();
        int fieldsLen = fields.length;

        // 遍历sheet中的列
        for (int columnNum = startColumn; columnNum <= sheet.getRow(rows[0] - 1).getLastCellNum(); columnNum++) {
            // 表头遇空时退出
            if (isEmptyCell(sheet.getRow(rows[0] - 1).getCell(columnNum - 1))) break;

            Object o;
            try {
                o = clazz.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            int fieldIndex = 0;
            for (int row : rows) {
                if (fieldIndex >= fieldsLen) break;
                // 忽略接收实体类中的字段，跳过数组长度辅助字段
                if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                    fieldIndex++;
                    continue;
                }
                setField(o, clazz, fields[fieldIndex], sheet.getRow(row - 1).getCell(columnNum - 1), sheetNo, row, columnNum);

                fieldIndex++;
            }

            // 重复性检查
            if (enableDuplicateCheck && list.size() > 0) {
                for (int i = 0; i < list.size(); i ++) {
                    if (o.equals(list.get(i))) {
                        throw new DataDuplicationException("文件工作表" + sheetNo + "第" + numberToAlphabet(i + startColumn) + "列与第" + numberToAlphabet(columnNum) + "列重复！");
                    }
                }
            }

            list.add(o);
        }

        return list;
    }


    /**
     * 获取数据垂直分布的表格的内容
     * 默认不进行重复性检查
     */
    public <T> List<T> getVerticalData(int sheetNo, int startRow, int startColumn, Class<T> clazz) throws IllegalStatementsException {
        return getVerticalData(sheetNo, startRow, startColumn, clazz, false);
    }


    /**
     * 获取数据垂直分布的表格的内容
     * @param sheetNo 第几个sheet
     * @param startRow 起始行
     * @param startColumn 起始列
     * @param clazz 接收实体类型
     * @param <T> 实体泛型
     * @return 目标sheet数据
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> getVerticalData(int sheetNo, int startRow, int startColumn, Class<T> clazz, boolean enableDuplicateCheck) throws IllegalStatementsException {
        if (sheetNo > sheets.length) {
            throw new SheetNoOutOfBoundsException("欲读取的工作表序号大于工作表总数！");
        }

        if (StringUtils.isNotEmpty(startTag) && startTag.equals(getComment(sheetNo, startRow, startColumn))) {
            throw new IllegalArgumentException("无法适配起始列！");
        }

        List<T> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        if (isBaseType(clazz)) {
            for (int columnNum = startColumn; columnNum <= sheet.getRow(startRow - 1).getLastCellNum(); columnNum++) {
                Row row = sheet.getRow(startRow - 1);
                Cell cell = row.getCell(columnNum - 1);

                if (isEmptyCell(cell)) break;

                String cellVal = getCellVal(cell);

                // 重复性检查
                if (enableDuplicateCheck && list.size() > 0) {
                    for (int i = 0; i < list.size(); i ++) {
                        if (cellVal.equals(list.get(i).toString())) {
                            throw new DataDuplicationException("文件工作表" + sheetNo + "第" + numberToAlphabet(i + startColumn) + "列与第" + numberToAlphabet(columnNum) + "列重复！");
                        }
                    }
                }

                list.add((T) ConvertUtils.convert(cellVal, clazz));
            }
        } else {
            // 遍历sheet中的列
            for (int columnNum = startColumn; columnNum <= sheet.getRow(startRow - 1).getLastCellNum(); columnNum++) {
                // 全空列时跳出
                if (isEmptyColumn(sheet, columnNum)) break;

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
                    // 忽略接收实体类中的字段，跳过数组长度辅助字段
                    if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                        fieldIndex++;
                        continue;
                    }
                    int assignedCellNum = setVerticalField(t, clazz, fields[fieldIndex], sheet.getRow(rowNum - 1).getCell(columnNum - 1), sheetNo, rowNum, columnNum);
                    rowNum += assignedCellNum;
                    fieldIndex ++;
                }

                // 重复性检查
                if (enableDuplicateCheck && list.size() > 0) {
                    for (int i = 0; i < list.size(); i ++) {
                        if (t.equals(list.get(i))) {
                            throw new DataDuplicationException("文件工作表" + sheetNo + "第" + numberToAlphabet(i + startColumn) + "列与第" + numberToAlphabet(columnNum) + "列重复！");
                        }
                    }
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
     * 实体类赋值
     *  包含注解防呆
     * @param o - 进行变量赋值的实体类
     * @param classType - 实体类的Class
     * @param field - 赋值字段
     * @param cell - 对应的Excel单元
     * @param rowNum - 行序号
     * @param columnNum - 列序号
     * @return 操作cell数
     */
    private int setVerticalField(Object o, Class<?> classType, Field field, Cell cell, int sheetNo, int rowNum, int columnNum) throws IllegalStatementsException {
        Object fieldValue;
        int assignedCell;

        if (field.isAnnotationPresent(DynamicRank.class)) {
            ParameterizedType types = (ParameterizedType) field.getGenericType();
            Type type = types.getActualTypeArguments()[0];

            DynamicRank dynamicRank = field.getAnnotation(DynamicRank.class);
            int[] columns = {dynamicRank.titleRank(), columnNum};
            boolean enableDuplicateCheck = dynamicRank.enableDuplicateCheck();

            List<Object> list = getVerticalDynamicRank(sheetNo, rowNum, columns, (Class<?>) type, enableDuplicateCheck);
            fieldValue = list;
            assignedCell = list.size();
        } else {
            // 获取通过注解校验的值
            fieldValue = getVerifiedCellVal(o, classType, field, cell, sheetNo, rowNum, columnNum);
            assignedCell = 1;
        }

        // 赋值
        setFieldValue(o, classType, field, fieldValue);

        return assignedCell;
    }


    /**
     * 获取垂直表格动态栏位
     * @param sheetNo sheet序号
     * @param startRow 开始行
     * @param columns 要获取动态栏位的列（必包括当前列，可包括表头栏等）
     * @param clazz 实体类类型
     * @param enableDuplicateCheck 重复性检查，依赖于实体类equals方法的实现
     * @return 动态栏位数据
     * @throws IllegalStatementsException 必解决异常
     */
    private List<Object> getVerticalDynamicRank(int sheetNo, int startRow, int[] columns, Class<?> clazz, boolean enableDuplicateCheck) throws IllegalStatementsException {
        List<Object> list = new ArrayList<>();
        Sheet sheet = sheets[sheetNo - 1];

        Field[] fields = clazz.getDeclaredFields();
        int fieldsLen = fields.length;

        for (int rowNum = startRow; rowNum <= sheet.getLastRowNum() + 1; rowNum++) {
            Row row = sheet.getRow(rowNum - 1);

            // 表头遇空时退出
            if (isEmptyCell(row.getCell(columns[0] - 1))) break;

            Object o;
            try {
                o = clazz.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            int fieldIndex = 0;
            for (int column : columns) {
                if (fieldIndex >= fieldsLen) break;
                // 忽略接收实体类中的字段，跳过数组长度辅助字段
                if (fields[fieldIndex].isAnnotationPresent(ValueIngnore.class)) {
                    fieldIndex++;
                    continue;
                }
                // 字段赋值
                setField(o, clazz, fields[fieldIndex], row.getCell(column - 1), sheetNo, rowNum, column);

                fieldIndex++;
            }

            // 重复性检查
            if (enableDuplicateCheck && list.size() > 0) {
                for (int i = 0; i < list.size(); i ++) {
                    if (o.equals(list.get(i))) {
                        throw new DataDuplicationException("文件工作表" + sheetNo + "第" + (i + startRow) + "行与第" + rowNum + "行重复！");
                    }
                }
            }

            list.add(o);
        }

        return list;
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
        // 获取通过注解校验的值
        Object fieldValue = getVerifiedCellVal(o, classType, field, cell, sheetNo, rowNum, columnNum);
        // 赋值
        setFieldValue(o, classType, field, fieldValue);
    }

    /**
     * 获取已通过注解防呆验证的String类型的Cell值
     * @param o - 操作的实体类
     * @param classType - 实体类的Class
     * @param field - 赋值字段
     * @param cell - 对应的Excel单元
     * @param rowNum - 行序号
     * @param columnNum - 列序号
     */
    private String getVerifiedCellVal(Object o, Class<?> classType, Field field, Cell cell, int sheetNo, int rowNum, int columnNum) throws IllegalStatementsException {
        String cellVal = getCellVal(cell, field);

        String columnAlphabet = numberToAlphabet(columnNum);
        // 非空
        if (field.isAnnotationPresent(NotNull.class)) {
            NotNull valuePattern = field.getAnnotation(NotNull.class);
            if (StringUtils.isEmpty(cellVal)) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnAlphabet + "列：" + valuePattern.message());
            }
        }

        // 格式限定
        if (field.isAnnotationPresent(ExcelDateFormat.class)) {
            ExcelDateFormat edf = field.getAnnotation(ExcelDateFormat.class);
            SimpleDateFormat sdf = new SimpleDateFormat(edf.pattern());
            try {
                sdf.parse(cellVal);
            } catch (ParseException e) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnAlphabet + "列：" + edf.message());
            }
        } else if (field.isAnnotationPresent(ExcelNumberFormat.class)) {
            if (StringUtils.isEmpty(cellVal)) {
                cellVal = "0";
            } else {
                ExcelNumberFormat enf = field.getAnnotation(ExcelNumberFormat.class);
                String when = enf.when();
                if (StringUtils.isEmpty(when) || isWhen(o, classType, when)) {
                    DecimalFormat df = new DecimalFormat(enf.format());
                    try {
                        df.format(new BigDecimal(cellVal));
                    } catch (Exception e) {
                        throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnAlphabet + "列：" + enf.message());
                    }
                }
            }
        } else if (field.isAnnotationPresent(ValueLimit.class)) {
            ValueLimit vl = field.getAnnotation(ValueLimit.class);
            String[] limit = vl.limit();
            Set<String> set = new HashSet<>(Arrays.asList(limit));
            if (!set.contains(cellVal)) {
                throw new IllegalArgumentException("文件工作表" + sheetNo + "第" + rowNum + "行第" + columnAlphabet + "列：" + vl.message());
            }
        }

        // 值转化
        if (field.isAnnotationPresent(Transform.class)) {
            Transform vl = field.getAnnotation(Transform.class);
            String[] expressions = vl.expressions();

            for (String expression : expressions) {
                if (!expression.contains("->")) {
                    throw new IllegalStatementsException("getexcelutil.annotation.Transform.expressions()：表达式“" + expression + "”缺失转化符号“->”");
                }

                String from = expression.substring(0, expression.indexOf("->")).trim();
                String to = expression.substring(expression.indexOf("->") + 2).trim();

                if (from.equals(cellVal)) {
                    cellVal = to;
                }
            }
        }

        return cellVal;
    }

    private boolean isWhen(Object target, Class<?> classType, String when) throws IllegalStatementsException {
        if (!when.contains("==")) {
            throw new IllegalStatementsException("getexcelutil.annotation.ExcelNumberFormat.when()：缺失条件连接符“==”");
        }

        String caseFieldName = when.substring(0, when.indexOf("==")).trim();
        String caseValue = when.substring(when.indexOf("==") + 2).trim();

        if (StringUtils.isEmpty(caseFieldName)) {
            throw new IllegalStatementsException("getexcelutil.annotation.ExcelNumberFormat.when()：条件判定字段不能为空！");
        }
        if (StringUtils.isEmpty(caseValue)) {
            throw new IllegalStatementsException("getexcelutil.annotation.ExcelNumberFormat.when()：条件判定值不能为空！");
        }

        String value;
        try {
            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(caseFieldName, classType);
            value = (String) propertyDescriptor.getReadMethod().invoke(target);
        } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
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
        try {
            PropertyDescriptor propertyDescriptor = new PropertyDescriptor(fieldName, classType);
            propertyDescriptor.getWriteMethod().invoke(target, values);
        } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
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

    private boolean isEmptyRow (Row row) {
        if (row != null) {
            for (Cell cell : row) {
                if (!isEmptyCell(cell)) {
                    return false;
                }
            }
        }
        return true;
    }

    private boolean isEmptyColumn (Sheet sheet, int columnNum) {
        for (int i = 0; i <= sheet.getLastRowNum(); i ++) {
            if (!isEmptyCell(sheet.getRow(i).getCell(columnNum - 1))) {
                return false;
            }
        }
        return true;
    }

    private boolean isEmptyCell (Cell cell) {
        return cell == null || cell.getCellType() == CellType.BLANK;
    }

    private boolean isBaseType (Class<?> clazz) {
        return String.class.equals(clazz)
                || Integer.class.equals(clazz)
                || Double.class.equals(clazz)
                || Object.class.equals(clazz)
                || Long.class.equals(clazz);
    }

    private String numberToAlphabet(int number) {
        if (number <= 0) {
            return null;
        }
        StringBuilder letter = new StringBuilder();
        do {
            -- number;
            int mod = number % 26; // 取余
            letter.append((char) (mod + 'A')); // 组装字符串
            number = (number - mod) / 26; // 计算剩下值
        } while (number > 0);
        return letter.reverse().toString(); // 反转
    }

}
