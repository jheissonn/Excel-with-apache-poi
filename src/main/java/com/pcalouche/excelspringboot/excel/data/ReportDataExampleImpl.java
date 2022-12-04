package com.pcalouche.excelspringboot.excel.data;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

public class ReportDataExampleImpl implements ReportData {

    private Iterator<ExampleData> iterator;

    @Override
    public void build() {
        iterator = Arrays.asList(
        new ExampleData(19, "Example 0", "Option - 2", 10D),
        new ExampleData(20, "Example 1", "Option - 0", 10D),
        new ExampleData(21, "Example 2", "Option - 1", 11D),
        new ExampleData(22, "Example 3", "Option - 1", 12D),
        new ExampleData(23, "Example 4", "Option - 2", 13D)).iterator();
    }

    @Override
    public boolean hasNext() {
        return iterator.hasNext();
    }

    @Override
    public List<Object> next() {
        if(iterator.hasNext()) {
            ExampleData exampleData = iterator.next();
            List<Object> row = new ArrayList<>();
            row.add(exampleData.getFieldExampleString());
            row.add(exampleData.getFieldExampleInteger());
            row.add(exampleData.getFieldExampleDropDown());
            row.add(exampleData.getFieldExampleValue());
            return row;
        }

        return null;
    }

    @Override
    public List<String> getTitles() {
        return Arrays.asList("fieldExampleInteger", "fieldExampleString", "fieldExampleDropDown", "fieldExampleValue");
    }

    @Override
    public List<DataValidationCustom> getValidations() {
        return Arrays.asList(new DataValidationCustom(2, new String[]{"Option - 0", "Option - 1", "Option - 2"}));
    }
}
