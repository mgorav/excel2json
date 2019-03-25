package com.gonnect.excel2json.model;

import lombok.Getter;
import lombok.Setter;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

@Getter
@Setter
public class AWorksheet implements Serializable {

    private String name;
    private List<ArrayList<Object>> data = new ArrayList<ArrayList<Object>>();
    private int maxCols = 0;

    public void addRow(ArrayList<Object> row) {
        data.add(row);
        if (maxCols < row.size()) {
            maxCols = row.size();
        }
    }

    public int getMaxRows() {
        return data.size();
    }

    public void fillColumns() {
        for (ArrayList<Object> tmp : data) {
            while (tmp.size() < maxCols) {
                tmp.add(null);
            }
        }
    }

}
