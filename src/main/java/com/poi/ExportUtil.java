package com.poi;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author YangQing
 * @version 1.0.0
 */

public abstract interface ExportUtil {
    abstract void setType(FileTypeEnum paramFileTypeEnum);

    abstract void setData(List<? extends Map> paramList);

    abstract void setData(LinkedHashMap<String, List<? extends Map>> paramLinkedHashMap);

    abstract void export();

    abstract void setHeaders(String[] paramArrayOfString);

    abstract void setHeaders(List<? extends String[]> paramList);

    abstract void setFields(String[] paramArrayOfString);

    abstract void setFields(List<? extends String[]> paramList);

    abstract void setTitle(String paramString);

    String exportUtil(String title, String[] headers, String[] fields, List<?> list);
}