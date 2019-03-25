package com.gonnect.excel2json.model;

import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.Getter;
import lombok.Setter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;

@Getter
@Setter
public class AWorkbook {

    private String fileName;
    private Collection<AWorksheet> sheets = new ArrayList<AWorksheet>();

    public void addExcelWorksheet(AWorksheet sheet) {
        sheets.add(sheet);
    }

    public String toJson() throws IOException {
        return new ObjectMapper().writer().withDefaultPrettyPrinter().writeValueAsString(this);
    }
}
