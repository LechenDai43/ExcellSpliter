package org.ExcellSpliter.ExcelProcessor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;

public class ExcelTable {
    private ArrayList<ArrayList<String>> content = new ArrayList<>();
    private ArrayList<String> tags = new ArrayList<>();
    private HashMap<String, Integer> tagIndex = new HashMap<>();
    private HashSet<String> pks;
    private int primaryKey = -1;

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
        ArrayList<String> keys = new ArrayList<>(),
        newCol = new ArrayList<>();
        for (int i = 0; i < et.tags.size(); i++) {
            if (tagIndex.containsKey(et.tags.get(i))) {
                keys.add(et.tags.get(i));
            } else {

            }
        }
    }
}
