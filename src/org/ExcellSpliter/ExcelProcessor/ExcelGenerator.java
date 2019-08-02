package org.ExcellSpliter.ExcelProcessor;

import java.io.File;

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


}
