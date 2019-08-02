package org.ExcellSpliter.ExcellPRocessor;

public class ExcelTester {
    private static void readerTester() {
        ExcelReader er = new ExcelReader("test1.xls");
        String[][] content = er.getSheet(0);
        for (String[] arr : content) {
            for (String str : arr) {
                if (str.length() > 0) {
                    System.out.print(str + "   ");
                } else {
                    System.out.print("?   ");
                }
            }
            System.out.println();
        }
    }

    public static void main(String[] args) {
        readerTester();
    }
}
