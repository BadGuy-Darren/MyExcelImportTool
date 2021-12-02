package com.darren.tools.getexcelutil.enums;

/**
 * 大迪
 * 实体类反射field类型String、hashCode
 */
public enum FieldType {

    NULL("null"),
    BOOLEAN("boolean"),
    BYTE("byte"),
    SHORT("short"),
    LONG("long"),
    INT("int"),
    FLOAT("float"),
    DOUBLE("double"),
    CHAR("char"),
    STRING("class java.lang.String"),
    INTEGER("class java.lang.Integer"),
    BIGDECIMAL("class java.math.BigDecimal"),
    BLONG("class java.lang.Long"),
    BSHORT("class java.lang.Short")
    ;

    private final String type;

    FieldType(String type) {
        this.type = type;
    }

    public static FieldType match(String type) {
        for(FieldType v:values()) {
            if(v.type.equals(type)) {
                return v;
            }
        }
        throw new IllegalArgumentException("Invalid Field Type: " + type);
    }

    public String getType() {
        return type;
    }

}
