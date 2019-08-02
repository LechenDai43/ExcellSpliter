package org.ExcellSpliter.ExcelProcessor;

import jxl.*;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;

/**
 * This class is for reading an Excel file.
 * @author Lechen Dai
 * @version 1.0
 * @since 8/1/2019
 */
public class ExcelReader {
    private Workbook workbook;
    private Sheet sheet;
    private int rowNum;
    private int colNum;

    /**
     * This constructor receives the address of the Excel file as a string.
     * If the string is null or is too short, then the work book is set to null.
     * If the file is not found, then the work book is set to null.
     * If the file is not excel file, then an exception is printed out.
     * @param address the address of the Excel file
     */
    public ExcelReader (String address) {
        if (address == null || address.length() < 4) {
            workbook = null;
        }
        try {
            workbook = Workbook.getWorkbook(new File(address));
        } catch (IOException e1) {
            workbook = null;
        } catch (BiffException e2) {
            e2.printStackTrace();
        }

    }

    /**
     * This constructor receives a file object.
     * If the file object is null, then the work book is set to null.
     * If the file object is not found, then the work book is set to null.
     * If the file object is not excel file, then an exception is printed out.
     * @param file the file that needs to be read as an excel
     */
    public ExcelReader (File file) {
        if (file == null) {
            workbook = null;
        }
        try {
            workbook = Workbook.getWorkbook(file);
        } catch (IOException e1) {
            workbook = null;
        } catch (BiffException e2) {
            e2.printStackTrace();
        }
    }

    /**
     * Get ready to read a specific sheet/page of the Excel file.
     * If the index is negative, then an illegal argument exception is thrown.
     * If the sheet does not exist in the file, then a null pointer exception is thrown.
     * If the sheet is empty, then a null pointer exception is thrown.
     * @param index the index of the sheet, starts from 0.
     * @throws IllegalArgumentException the index cannot be negative
     * @throws NullPointerException the sheet must exist and contain at least one element
     */
    public void readSheet (int index) {
        if (index < 0) {
            throw new IllegalArgumentException("Index cannot be negative.");
        }
        sheet = workbook.getSheet(index);
        if (sheet == null || sheet.getRows() == 0 || sheet.getColumns() == 0) {
            throw new NullPointerException("This sheet does not exist or is empty.");
        }
        colNum = 0;
        rowNum = 0;
    }

    /**
     * Set the reading sheet to the specified index, and return how many rows in the sheet.
     * It may throw illegal argument exception or null pointer exception.
     * @param index the number of the sheet you want to read
     * @return the number of rows in the specified sheet
     * @see this#readSheet(int)
     * @see this#getRowNum()
     */
    public int getRowNum (int index) {
        readSheet(index);
        return getRowNum();
    }

    /**
     * Get how many rows in the current sheet.
     * @return the number of rows in the current sheet
     * @throws NullPointerException no sheet is ready to read
     * @see this#getRowNum(int)
     */
    public int getRowNum () {
        if (sheet != null) {
            return sheet.getRows();
        } else {
            throw new NullPointerException("No sheet is ready to read.");
        }
    }

    /**
     * Set the reading sheet to the specified index, and return how many columns in the sheet.
     * It may throw illegal argument exception or null pointer exception.
     * @param index the number of sheet to read
     * @return the number of columns in the specified sheet
     * @see this#readSheet(int)
     * @see this#getColNum()
     */
    public int getColNum (int index) {
        readSheet(index);
        return getColNum();
    }

    /**
     * Get how many columns in the current sheet.
     * @return the number of columns of the current sheet.
     * @throws NullPointerException no sheet is ready to read
     * @see this#getColNum(int)
     */
    public int getColNum () {
        if (sheet != null) {
            return sheet.getColumns();
        } else {
            throw new NullPointerException("No sheet is ready to read.");
        }
    }

    /**
     * Get the content of a whole column.
     * @param index the index of the column that needs to be read
     * @return a string array that each string contains all content in a cell
     * @throws NullPointerException no sheet is ready to read
     * @throws IllegalArgumentException the index cannot be negative
     * @throws IndexOutOfBoundsException specified an empty column
     * @see this#getCol(int, int)
     * @see this#getCol()
     */
    public String[] getCol (int index) {
        if (sheet == null) {
            throw new NullPointerException("No sheet is ready for reading.");
        }
        if (index < 0) {
            throw new IllegalArgumentException("The index cannot be negative.");
        }
        if (index > getColNum()) {
            throw new IndexOutOfBoundsException("There is no enough columns in the current sheet.");
        }
        Cell[] cells = sheet.getColumn(index);
        String[] result = new String[cells.length];
        for (int i = 0; i < cells.length; i++) {
            result[i] = cells[i].getContents();
        }
        colNum = index + 1;
        return result;
    }

    /**
     * Get a whole specified column in a specified sheet page.
     * @param sheetNum the number of the specified sheet
     * @param columnNum the number of the specified column
     * @return a string array that each string is the content of a cell
     * @see this#readSheet(int)
     * @see this#getCol(int)
     */
    public String[] getCol (int sheetNum, int columnNum) {
        readSheet(sheetNum);
        return getCol(columnNum);
    }

    /**
     * Get the whole next column.
     * @return a string array that each string is the content of a cell
     * @see this#getCol(int)
     */
    public String[] getCol () {
        return getCol(colNum);
    }

    /**
     * Get the index of the column that will be processed by the next call of getCol().
     * @return the index of the column
     */
    public int getCurrentColNum () {
        return colNum;
    }

    /**
     * Get the content of a whole row.
     * @param index the index of the row that needs to be read
     * @return a string array that each string contains all content in a cell
     * @throws NullPointerException no sheet is ready to read
     * @throws IllegalArgumentException the index cannot be negative
     * @throws IndexOutOfBoundsException specified an empty row
     * @see this#getRow(int, int)
     * @see this#getRow()
     */
    public String[] getRow (int index) {
        if (sheet == null) {
            throw new NullPointerException("No sheet is ready for reading.");
        }
        if (index < 0) {
            throw new IllegalArgumentException("The index cannot be negative.");
        }
        if (index > getRowNum()) {
            throw new IndexOutOfBoundsException("There is no enough rows in the current sheet.");
        }
        Cell[] cells = sheet.getRow(index);
        String[] result = new String[cells.length];
        for (int i = 0; i < cells.length; i++) {
            result[i] = cells[i].getContents();
        }
        rowNum = index + 1;
        return result;
    }

    /**
     * Get a whole specified row in a specified sheet page.
     * @param sheetNum the number of the specified sheet
     * @param rowNum the number of the specified row
     * @return a string array that each string is the content of a cell
     * @see this#readSheet(int)
     * @see this#getRow(int)
     */
    public String[] getRow (int sheetNum, int rowNum) {
        readSheet(sheetNum);
        return getCol(rowNum);
    }

    /**
     * Get the whole next row.
     * @return a string array that each string is the content of a cell
     * @see this#getRow(int)
     */
    public String[] getRow () {
        return getRow(rowNum);
    }

    /**
     * Get the index of the row that will be processed by the next call of getRow().
     * @return the index of the row
     * @see this#getRow()
     */
    public int getCurrentRowNum () {
        return rowNum;
    }

    /**
     * Get all content in a certain range in a 2D string array.
     * @param x1 index of starting column
     * @param x2 index of ending column
     * @param y1 index of starting row
     * @param y2 index of ending row
     * @return the 2D string array contains the content of the given range
     * @throws NullPointerException no sheet is ready to read
     * @throws IllegalArgumentException the parameters themselves have logic issue
     * @throws IndexOutOfBoundsException any index points to an empty row or an empty column
     */
    public String[][] getSheet (int x1, int x2, int y1, int y2) {
        if (sheet == null) {
            throw new NullPointerException("No sheet is ready to read.");
        }
        if (x1 < 0 || x2 < 0 || y1 < 0 || y2 < 0) {
            throw new IllegalArgumentException("No parameter can be zero.");
        } else if (x1 > x2) {
            throw new IllegalArgumentException("Starting column index cannot larger than ending column index.");
        } else if (y1 > y2) {
            throw new IllegalArgumentException("Starting row index cannot larger than ending row index.");
        }
        if (x1 >= sheet.getColumns() || x2 >= sheet.getColumns()) {
            throw new IndexOutOfBoundsException("Column indices exceed the number of column.");
        } else if (y1 >= sheet.getRows() || y2 >= sheet.getRows()) {
            throw new IndexOutOfBoundsException("Row indices exceed the number of row.");
        }
        String[][] result = new String[y2 - y1 + 1][x2 - x1 + 1];
        for (int i = 0; i < result.length; i++) {
            String[] row = this.getRow(y1 + i);
            for (int j = 0; j < result[i].length; j++) {
                if (x1 + j < row.length) {
                    result[i][j] = row[x1 + j];
                } else {
                    result[i][j] = "";
                }
            }
        }
        return result;
    }

    /**
     * Get the whole content of the current sheet.
     * @return the 2D string array contains the content of the whole sheet
     * @see this#getSheet(int, int, int, int)
     */
    public String[][] getSheet () {
        return getSheet(0, sheet.getColumns() - 1, 0, sheet.getRows() - 1);
    }

    /**
     * Get the whole content of the specified sheet.
     * @param index the index of the sheet
     * @return the 2D string array contains the content of the specified page
     * @see this#getSheet(int, int, int, int)
     * @see this#getSheet()
     * @see this#readSheet(int)
     */
    public String[][] getSheet (int index) {
        readSheet(index);
        return getSheet();
    }

    /**
     * Get the content in a given range of a specified sheet.
     * @param index the index of the sheet
     * @param x1 the starting column index
     * @param x2 the ending column index
     * @param y1 the starting row index
     * @param y2 the ending row index
     * @return the 2D string array that contains the content of the given range in the specified sheet
     * @see this#readSheet(int)
     * @see this#getSheet(int, int, int, int)
     */
    public String [][] getSheet (int index, int x1, int x2, int y1, int y2) {
        readSheet(index);
        return getSheet(x1, x2, y1, y2);
    }

    /**
     * Free the memory that is allocated to the work book.
     */
    public void close () {
        sheet = null;
        if (workbook != null) {
            workbook.close();
        }
    }

    /**
     * Close the old excel file and open a new one by the given address.
     * @param address the address of the new excel file
     */
    public void open (String address) {
        if (workbook != null) {
            workbook.close();
        }
        if (address == null || address.length() < 4) {
            workbook = null;
        }
        try {
            workbook = Workbook.getWorkbook(new File(address));
        } catch (IOException e1) {
            workbook = null;
        } catch (BiffException e2) {
            e2.printStackTrace();
        }
    }

    /**
     * Close the old excel file and open a new one that is specified by a file object.
     * @param file the new excel file
     */
    public void open (File file) {
        if (workbook != null) {
            workbook.close();
        }
        if (file == null) {
            workbook = null;
        }
        try {
            workbook = Workbook.getWorkbook(file);
        } catch (IOException e1) {
            workbook = null;
        } catch (BiffException e2) {
            e2.printStackTrace();
        }
    }
}
