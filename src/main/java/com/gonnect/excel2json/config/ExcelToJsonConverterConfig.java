package com.gonnect.excel2json.config;

import lombok.Getter;
import lombok.Setter;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;


@Getter
@Setter
public class ExcelToJsonConverterConfig {

    private String sourceFile;
    private boolean parsePercentAsFloats = false;
    private boolean omitEmpty = false;
    private boolean pretty = false;
    private boolean fillColumns = false;
    private int numberOfSheets = 0;
    private int rowLimit = 0; // 0 -> no limit
    private int rowOffset = 0;
    private DateFormat formatDate = null;



    public static ExcelToJsonConverterConfig create(ExcelToJsonParsingOptions cmd) {
        ExcelToJsonConverterConfig config = new ExcelToJsonConverterConfig();

        if(cmd.hasOption("s")) {
            config.sourceFile = cmd.getOptionValue("s");
        }

        if(cmd.hasOption("df")) {
            config.formatDate = new SimpleDateFormat(cmd.getOptionValue("df"));
        }
        if (cmd.hasOption("n")) {
            config.setNumberOfSheets(Integer.valueOf(cmd.getOptionValue("n")));
        }

        if (cmd.hasOption("l")) {
            config.setRowLimit(Integer.valueOf(cmd.getOptionValue("l")));
        }
        if (cmd.hasOption("o")) {
            config.setRowOffset(Integer.valueOf(cmd.getOptionValue("o")));
        }

        config.parsePercentAsFloats = cmd.hasOption("percent");
        config.omitEmpty = !cmd.hasOption("empty");
        config.pretty = cmd.hasOption("pretty");
        config.fillColumns = cmd.hasOption("fillColumns");


        return config;
    }

    public String valid() {
        if(sourceFile==null) {
            return "Source file may not be empty.";
        }
        File file = new File(sourceFile);
        if(!file.exists()) {
            return "Source file does not exist.";
        }
        if(!file.canRead()) {
            return "Source file is not readable.";
        }

        return null;
    }


}