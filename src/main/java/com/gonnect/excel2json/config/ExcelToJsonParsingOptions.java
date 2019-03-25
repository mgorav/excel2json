package com.gonnect.excel2json.config;

import java.util.HashMap;
import java.util.Map;

/**
    s,source <arg>     The source file which should be converted into json.
    df,--dateFormat    The template to use for fomatting dates into strings.
    l,rowLimit <arg>   Limit the max number of rows to read.
    n,maxSheets <arg>  Limit the max number of sheets to read.
    o,rowOffset <arg>  Set the offset for begin to read.
    empty              Include rows with no data in it.
    percent            Parse percent values as floats.
 */
public class ExcelToJsonParsingOptions {

    private Map<String, String> options = new HashMap<>();


    public boolean hasOption(String param) {

        return options.containsKey(param);
    }

    public void addOption(String option, String value) {
        options.put(option, value);
    }

    public String getOptionValue(String param) {

        return options.get(param);
    }
}
