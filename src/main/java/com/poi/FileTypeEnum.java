package com.poi;

import java.util.Arrays;
import java.util.Iterator;

/**
 * @author YangQing
 * @version 1.0.0
 */
public enum FileTypeEnum
{
    EXCEL, CVS, GZIP, JSON, TXT, ZIP;

    private String name;

    public String getName()
    {
        return this.name; }

    public static FileTypeEnum byName(String name) {
        for (Iterator localIterator = Arrays.asList(values()).iterator();
             localIterator.hasNext(); ) {
            FileTypeEnum ftype = (FileTypeEnum)localIterator.next();
            if (ftype.getName().equals(name))
                return ftype;
        }
        return null;
    }
}