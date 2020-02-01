package com.poiji.bind.mapping;

import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelCellName;
import com.poiji.annotation.ExcelCellRange;
import com.poiji.annotation.ExcelRow;
import com.poiji.annotation.ExcelUnknownCells;
import com.poiji.config.Casting;
import com.poiji.exception.IllegalCastException;
import com.poiji.exception.InvalidModelException;
import com.poiji.option.PoijiOptions;
import com.poiji.util.AnnotationUtil;
import com.poiji.util.ReflectUtil;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Optional;
import java.util.function.Consumer;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import static com.poiji.util.ReflectUtil.getAllFields;
import static java.lang.String.valueOf;

/**
 * This class handles the processing of a .xlsx file,
 * and generates a list of instances of a given type
 * <p>
 * Created by hakan on 22/10/2017
 */
final class PoijiHandler<T> implements SheetContentsHandler {
    private T instance;
    private Consumer<? super T> consumer;
    private int internalRow;
    private int internalCount;
    private int limit;
    private Class<T> type;
    private PoijiOptions options;
    private final Casting casting;
    private Map<String, Integer> columnIndexPerTitle;
    private Map<Integer, String> titlePerColumnIndex;
    // New maps used to speed up computing and handle inner objects
    private Map<String, Object> fieldInstances;
    private Map<Object, Optional<Field>> columnToField;
    private Map<Integer, Field> columnToSuperClassField;
    private Set<ExcelCellName> excelCellNames;

    PoijiHandler(Class<T> type, PoijiOptions options, Consumer<? super T> consumer) {
        this.type = type;
        this.options = options;
        this.consumer = consumer;
        this.limit = options.getLimit();
        casting = options.getCasting();
        columnIndexPerTitle = new HashMap<>();
        titlePerColumnIndex = new HashMap<>();
        columnToField = new HashMap<>();
        columnToSuperClassField = new HashMap<>();
        excelCellNames = new HashSet<>();

        parseTypeAnnotations(type);
    }

    private void setFieldValue(String content, Class<? super T> subclass, int column) {
        setValue(content, subclass, column);
    }

    /**
     * Using this to hold inner objects that will be mapped to the main object
     **/
    private Object getInstance(Field field) {
        Object ins = null;
        if (fieldInstances.containsKey(field.getName())) {
            ins = fieldInstances.get(field.getName());
        } else {
            ins = ReflectUtil.newInstanceOf(field.getType());
            fieldInstances.put(field.getName(), ins);
        }
        return ins;
    }


    private void ensureExclusiveAnnotationPerField(Field field, Class<? extends Annotation>... annotationTypes) {
        List<Class<? extends Annotation>> presentAnnotations =
                Arrays.stream(annotationTypes)
                        .filter(annotationType -> field.getAnnotation(annotationType) != null)
                        .collect(Collectors.toList());

        if (presentAnnotations.size() > 1) {
            throw new InvalidModelException("Field " + field.getName() + " has annotations that are not allowed to be mixed: " + presentAnnotations);
        }
    }

    private void parseTypeAnnotations(Class<?> type) {
        List<Field> fields = getAllFields(type);

        fields.forEach(field -> ensureExclusiveAnnotationPerField(field,
                ExcelRow.class, ExcelCell.class, ExcelCellName.class, ExcelCellRange.class, ExcelUnknownCells.class));

        // Check presence of @ExcelRange
        fields.stream()
                .filter(field -> field.getAnnotation(ExcelCellRange.class) != null)
                .forEach(field -> parseTypeAnnotations(field.getType()));

        // Check presence of @ExcelRow
        List<Field> excelRowAnnotated = fields.stream()
                .filter(field -> field.getAnnotation(ExcelRow.class) != null)
                .collect(Collectors.toList());

        if (excelRowAnnotated.size() > 1) {
            throw new InvalidModelException("@ExcelRow annotation used more than once: " + excelRowAnnotated);
        } else if (excelRowAnnotated.size() == 1) {
            columnToField.put(ExcelRow.class, Optional.of(excelRowAnnotated.get(0)));
        }

        // Check presence of @ExcelCell
        fields.stream()
                .filter(field -> field.getAnnotation(ExcelCell.class) != null)
                .forEach(field -> columnToField.put(field.getAnnotation(ExcelCell.class).value(), Optional.of(field)));

        // Check presence of @ExcelCellName
        fields.stream()
                .filter(field -> field.getAnnotation(ExcelCellName.class) != null)
                .forEach(field -> {
                    columnToField.put(field.getAnnotation(ExcelCellName.class).value(), Optional.of(field));
                });

        // Check presence of @ExcelUnknownCells
        List<Field> excelUnknownCellsAnnotated = fields.stream()
                .filter(field -> field.getAnnotation(ExcelUnknownCells.class) != null)
                .collect(Collectors.toList());

        if (excelUnknownCellsAnnotated.size() > 1) {
            throw new InvalidModelException("@ExcelUnknownCells annotation used more than once: " + excelUnknownCellsAnnotated);
        } else if (excelUnknownCellsAnnotated.size() == 1) {
            // assert type of Map<String, String>
            Field field = excelUnknownCellsAnnotated.get(0);
            if (field.getType() != Map.class) {
                throw new InvalidModelException("@ExcelUnknownCells can only be used on Map<String,String> types");
            }
            Type[] actualTypeArguments = ((ParameterizedType) field.getGenericType()).getActualTypeArguments();
            if (actualTypeArguments[0] != String.class || actualTypeArguments[1] != String.class) {
                throw new InvalidModelException("@ExcelUnknownCells can only be used on Map<String,String> types");
            }

            columnToField.put(ExcelUnknownCells.class, Optional.of(field));
        }
    }

    /**
     * Adds content to the unknown cells map
     *
     * @param field      of the Map<String,String> where to put the data
     * @param content    will be put to value of the map
     * @param columnName will be used as key of the map
     */
    private void addToUnknownCellsMap(Field field, String content, String columnName) {
        try {
            Map<String, String> excelUnknownCellsMap;
            field.setAccessible(true);
            if (field.get(instance) == null) {
                excelUnknownCellsMap = new HashMap<>();
                setFieldData(field, excelUnknownCellsMap, instance);
            } else {
                excelUnknownCellsMap = (Map) field.get(instance);
            }

            excelUnknownCellsMap.put(columnName, content);
        } catch (IllegalAccessException e) {
            throw new IllegalCastException("Could not read content of field " + field.getName() + " on Object {" + instance + "}");
        }
    }

    private void setExcelRowIndex() {
        columnToField.getOrDefault(ExcelRow.class, Optional.empty())
                .ifPresent(field -> setFieldData(field, internalRow, instance));
    }

    private boolean setValue(String content, Class<? super T> type, int column) {
        Optional<Field> mappedFieldOptional = columnToField.getOrDefault(column, Optional.empty());

        if(!mappedFieldOptional.isPresent()) {
            String columnName = titlePerColumnIndex.get(column);
            mappedFieldOptional = columnToField.getOrDefault(columnName, Optional.empty());
        }

        // TODO: Don't do this more than once
        setExcelRowIndex();


        if (mappedFieldOptional.isPresent()) {
            Field mappedField = mappedFieldOptional.get();

            setValue(mappedField, column, content, instance);

//            if (mappedField.getAnnotation(ExcelCellRange.class) != null) {
//                Object ins = getInstance(mappedField);
//            }
        } else {
            columnToField.getOrDefault(ExcelUnknownCells.class, Optional.empty())
                    .ifPresent(field -> {
                        String columnName = titlePerColumnIndex.get(column);
                        addToUnknownCellsMap(field, content, columnName);
                    });
        }

        Stream.of(type.getDeclaredFields())
                .filter(field -> field.getAnnotation(ExcelUnknownCells.class) == null)
                .forEach(field -> {
                    ExcelCellRange range = field.getAnnotation(ExcelCellRange.class);
                    if (range != null) {
                        Object ins = null;
                        ins = getInstance(field);
                        for (Field f : field.getType().getDeclaredFields()) {
                            if (setValue(f, column, content, ins)) {
                                setFieldData(field, ins, instance);
                                columnToField.put(column, Optional.of(f));
                                columnToSuperClassField.put(column, field);
                            }
                        }
                    } else {
                        if (setValue(field, column, content, instance)) {
                            columnToField.put(column, Optional.of(field));
                        }
                    }
                });


        return false;
    }

    private boolean setValue(Field field, int column, String content, Object ins) {
        ExcelCell index = field.getAnnotation(ExcelCell.class);
        if (index != null) {
            Class<?> fieldType = field.getType();
            if (column == index.value()) {
                Object o = casting.castValue(fieldType, content, internalRow, column, options);
                setFieldData(field, o, ins);
                return true;
            }
        } else {
            ExcelCellName excelCellName = field.getAnnotation(ExcelCellName.class);
            if (excelCellName != null) {
                excelCellNames.add(excelCellName);
                Class<?> fieldType = field.getType();
                final String titleName = options.getCaseInsensitive()
                        ? excelCellName.value().toLowerCase()
                        : excelCellName.value();
                final Integer titleColumn = columnIndexPerTitle.get(titleName);
                //Fix both columns mapped to name passing this condition below
                if (titleColumn != null && titleColumn == column) {
                    Object o = casting.castValue(fieldType, content, internalRow, column, options);
                    setFieldData(field, o, ins);
                    return true;
                }
            }
        }
        return false;
    }

    private void setFieldData(Field field, Object o, Object instance) {
        try {
            field.setAccessible(true);
            field.set(instance, o);
        } catch (IllegalAccessException e) {
            throw new IllegalCastException("Unexpected cast type {" + o + "} of field" + field.getName());
        }
    }

    @Override
    public void startRow(int rowNum) {
        if (rowNum + 1 > options.skip()) {
            internalCount += 1;
            instance = ReflectUtil.newInstanceOf(type);
            fieldInstances = new HashMap<>();
        }
    }

    @Override
    public void endRow(int rowNum) {
        if (internalRow != rowNum)
            return;

        if (rowNum + 1 > options.skip()) {
            consumer.accept(instance);
        }
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        CellAddress cellAddress = new CellAddress(cellReference);
        int row = cellAddress.getRow();
        int headers = options.getHeaderStart();
        int column = cellAddress.getColumn();
        if (row <= headers) {
            columnIndexPerTitle.put(
                    options.getCaseInsensitive() ? formattedValue.toLowerCase() : formattedValue,
                    column
            );
            titlePerColumnIndex.put(column, getTitleNameForMap(formattedValue, column));
        }
        if (row + 1 <= options.skip()) {
            return;
        }
        if (limit != 0 && internalCount > limit) {
            return;
        }
        internalRow = row;
        setFieldValue(formattedValue, type, column);
    }

    private String getTitleNameForMap(String cellContent, int columnIndex) {
        String titleName;
        if (titlePerColumnIndex.containsValue(cellContent)
                || cellContent.isEmpty()) {
            titleName = cellContent + "@" + columnIndex;
        } else {
            titleName = cellContent;
        }
        return titleName;
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        //no-op
    }

    @Override
    public void endSheet() {
        AnnotationUtil.validateMandatoryNameColumns(options, type, columnIndexPerTitle.keySet());
    }
}
