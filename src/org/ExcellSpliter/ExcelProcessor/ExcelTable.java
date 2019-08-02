package org.ExcellSpliter.ExcelProcessor;

import jxl.Cell;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCell;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

public class ExcelTable {
    private ArrayList<ArrayList<String>> content = new ArrayList<>();
    private ArrayList<String> tags = new ArrayList<>();
    private HashMap<String, Integer> tagIndex = new HashMap<>();

    public ExcelTable () {

    }

    public ExcelTable (ArrayList<String> head) {
        if (head != null) {
            tags = head;
            for (int i = 0; i < tags.size(); i++) {
                tagIndex.put(tags.get(i), i);
            }
        }
    }

    /*
    Add some more constructor later.
     */

    public int[] getSize () {
        return new int[]{content.size(), tags.size()};
    }

    public boolean setPK (int keyCol) {
        if (keyCol < 0) {
            throw new IllegalArgumentException("Primary column index cannot be negative.");
        }
        if (keyCol > tags.size()) {
            throw new IllegalArgumentException("Primary column index cannot larger than the number of columns.");
        }
        boolean seccess = true;

        return seccess;
    }

    public void addColumns (ArrayList<String> moreTags) {
        if (moreTags == null) {
            throw new NullPointerException("The input array list is null.");
        }
        for (int i = 0; i < moreTags.size(); i++) {
            if (tagIndex.containsKey(moreTags.get(i))) {
                continue;
            } else {
                tags.add(moreTags.get(i));
                for (int j = 0; j < content.size(); j++) {
                    content.get(j).add("");
                }
                tagIndex.put(moreTags.get(i), tags.size() - 1);
            }
        }
    }

    public void addRows(int num) {
        if (num < 0) {
            throw new IllegalArgumentException("Cannot insert negative number of rows");
        }
        for (int i = 0; i < num; i++) {
            ArrayList<String> row = new ArrayList<>();
            for (int j = 0; j < tags.size(); j++) {
                row.add("");
            }
            content.add(row);
        }
    }

    public void addTable (ExcelTable et) {
        //Classify which tag is in both tables and which is not.
        //Use the tags that are in both tables as unique key.
        //The tags that are only in the et table will be insert into this table later.
        ArrayList<String> keys = new ArrayList<>(),
        newCol = new ArrayList<>();
        for (int i = 0; i < et.tags.size(); i++) {
            if (tagIndex.containsKey(et.tags.get(i))) {
                keys.add(et.tags.get(i));
            } else {
                newCol.add(et.tags.get(i));
            }
        }
        //Insert the columns that are only in et table into this table.
        this.addColumns(newCol);
        //Create an index for unique keys in this table.
        HashMap<ArrayList<String>, Integer> uks = new HashMap<>();
        for (int i = 0; i < content.size(); i++) {
            ArrayList<String> k = new ArrayList<>();
            for (int j = 0; j < keys.size(); j++) {
                String att = content.get(i).get(tagIndex.get(keys.get(j)));
                k.add(att);
            }
            if (uks.containsKey(k)) {
                continue;
            } else {
                uks.put(k, i);
            }
        }
        //Check each row of the et table to see if its unique key is already in this table.
        //If the key exists, then update the row.
        //If not, then insert a new row and add values to it.
        for (int i = 0; i < et.content.size(); i++) {
            //get the unique key of this row in et table.
            ArrayList<String> k = new ArrayList<>();
            for (int j = 0; j < keys.size(); j++) {
                String att = et.content.get(i).get(et.tagIndex.get(keys.get(j)));
                k.add(att);
            }
            //check if this row is in this table and record its index
            int rowIndex;
            if (uks.containsKey(k)) {
                rowIndex = uks.get(k);
            } else {
                rowIndex = this.content.size();
                this.addRows(1);
            }
            //put the value from et table to this table
            for (int j = 0; j < et.tags.size(); j++) {
                String value = et.content.get(i).get(et.tagIndex.get(et.tags.get(j)));
                int index = this.tagIndex.get(et.tags.get(j));
                this.content.get(rowIndex).set(index, value);
            }
        }
    }

    public WritableCell[][] toExcelSheet (HashMap<Integer, String> condition) {
        WritableCell[][] result = new WritableCell[content.size() + 1][tags.size()];
        for (int i = 0; i < tags.size(); i++) {
            WritableCell t = new Label(i, 0, tags.get(i));
            result[0][i] = t;
        }
        for (int i = 0; i < content.size(); i++) {
            ArrayList<String> row = content.get(i);
            for (int j = 0; j < row.size(); j++) {
                WritableCell c;
                if (condition.containsKey(j)) {
                    String sk = condition.get(j);
                    switch (sk) {
                        case "num": case "Num": case "number": case "Number":
                            if (row.get(j).equals("")) {
                                c = new Number(j, i + 1, 0);
                            } else {
                                double doubleValue = Double.parseDouble(row.get(j));
                                c = new Number(j, i + 1, doubleValue);
                            }
                            break;
                        case "date": case "Date":
                            if (row.get(j).equals("")) {
                                c = new DateTime(j, i + 1, null);
                            } else {
                                Date dateValue = new Date(row.get(j));
                                c = new DateTime(j, i + 1, dateValue);
                            }
                            break;
                        default:
                            c = new Label(j, i + 1, row.get(j));
                            break;
                    }
                } else {
                    c = new Label(j, i + 1, row.get(j));
                }
                result[i + 1][j] = c;
            }
        }
        return result;
    }

    public WritableCell[][] toExcelSheet () {
        return toExcelSheet(new HashMap<Integer, String>());
    }

    public void addCell (int r, int c, String str) {
        if (r >= content.size() || c >= tags.size()) {
            System.out.println(r + " " + content.size() + "," + c + " " + tags.size());
            throw new IndexOutOfBoundsException("Row index or column index is too large.");
        }
        content.get(r).set(c, str);
    }

    public void addCell (int r, String t, String str) {
        if (tagIndex.containsKey(t)) {
            addCell(r, tagIndex.get(t), str);
        } else {
            throw new IllegalArgumentException("Do not have such tag.");
        }
    }
}
