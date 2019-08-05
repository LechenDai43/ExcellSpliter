package org.ExcellSplitter.ExcelProcessor;

import jxl.Workbook;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelGenerator {

    private File file;

    public ExcelGenerator (String address) {
        if (address.endsWith(".xls")) {
            file = new File(address);
            if (file.exists()) {
                file = null;
            }
        } else {
            file = null;
        }
    }

    public boolean locateFile (String address) {
        if (address.endsWith(".xls")) {
            File old = file;
            file = new File(address);
            if (file.exists()) {
                file = old;
                return false;
            }
        } else {
            return false;
        }
        return true;
    }

    public void createFile (ArrayList<ExcelTable> tables) throws IOException, WriteException {
        if (tables == null || tables.size() < 1) {
            throw new IllegalArgumentException("There must be at least one table element.");
        }
        if (!file.exists()) {
            file.createNewFile();
        }
        WritableWorkbook workbook = Workbook.createWorkbook(file);
        for (int i = 0; i < tables.size(); i++) {
            WritableSheet sheet = workbook.createSheet("Sheet " + (i + 1), i);
            WritableCell[][] cells = tables.get(i).toExcelSheet();
            for (int j = 0; j < cells.length; j++) {
                for (int k = 0; k < cells[j].length; k++) {
                    sheet.addCell(cells[j][k]);
                }
            }
        }
        workbook.write();
        workbook.close();
        file = null;
    }
}
