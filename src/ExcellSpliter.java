import jxl.write.WriteException;
import org.ExcellSplitter.CSVFileOrganizer.CSVOrganizer;
import org.ExcellSplitter.ExcelProcessor.ExcelGenerator;
import org.ExcellSplitter.ExcelProcessor.ExcelReader;
import org.ExcellSplitter.ExcelProcessor.ExcelTable;
import org.ExcellSplitter.GUI.ESFrame;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.StringTokenizer;

public class ExcellSpliter {

    public static ArrayList<ArrayList<String>> spliteCell (String[][] content, int index, String delimiter) {
        HashMap<String, Integer> tags = new HashMap<>();
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        result.add(new ArrayList<String>());
        String root = content[0][index] + "-";
        int pos = 0;
        for (int i = 1; i < content.length; i++) {
            String whole = content[i][index];
            StringTokenizer st = new StringTokenizer(whole, delimiter);
            HashMap<Integer, String> rowIndex = new HashMap<>();
            while (st.hasMoreTokens()) {
                String token = st.nextToken();
                if (tags.containsKey(token)) {
                    rowIndex.put(tags.get(token), token);
                } else {
                    tags.put(token, pos);
                    rowIndex.put(pos, token);
                    pos++;
                    result.get(0).add(root + token);
                }
            }
            ArrayList<String> row = new ArrayList<>();
            for (int j = 0; j < result.get(0).size(); j++) {
                if (rowIndex.containsKey(j)) {
                    row.add(rowIndex.get(j));
                } else {
                    row.add("");
                }
            }
            result.add(row);
        }
        for (int i = 1; i < result.size(); i++) {
            if (result.get(i).size() < result.get(0).size()) {
                while (result.get(i).size() < result.get(0).size()) {
                    result.get(i).add("");
                }
            }
        }
        return result;
    }

    public static ArrayList<ArrayList<String>> spliteCells (String[][] content, int[] indices, String delimiter) {
        if (indices == null || indices.length < 1) {
            throw new IllegalArgumentException("There must be at least one index.");
        }
        if (content == null) {
            throw new NullPointerException("The content cannot be null.");
        }
        ArrayList<ArrayList<String>> result = spliteCell(content, indices[0], delimiter);
        for (int i = 1; i < indices.length; i++) {
            ArrayList<ArrayList<String>> tem = spliteCell(content, indices[i], delimiter);
            for (int j = 0; j < result.size(); j++) {
                result.get(j).addAll(tem.get(j));
            }
        }
        return result;
    }

    public static ExcelTable toTable (ArrayList<ArrayList<String>> arr, String[][] content, int[] indices) {
        ExcelTable table = new ExcelTable();
        HashSet<Integer> indexSet = new HashSet<>();
        for (int i = 0; i < indices.length; i++) {
            indexSet.add(indices[i]);
        }
        ArrayList<String> basicTags = new ArrayList<>();
        for (int i = 0; i < content[0].length; i++) {
            if (!indexSet.contains(i)) {
                basicTags.add(content[0][i]);
            }
        }
        table.addColumns(basicTags);
        table.addRows(content.length - 1);
        table.addColumns(arr.get(0));
        for (int i = 1; i < content.length; i++) {
            int colNum = 0;
            for (int j = 0; j < content[i].length; j++) {
                if (!indexSet.contains(j)) {
                    table.addCell(i - 1, colNum, content[i][j]);
                    colNum++;
                }
            }
            for (int j = 0; j < arr.get(i).size(); j++) {
                table.addCell(i - 1, colNum, arr.get(i).get(j));
                colNum++;
            }
        }
        return table;
    }

    public static void main (String[] args) throws IOException, WriteException {
//        ExcelReader er = new ExcelReader("input.xls");
        ExcelGenerator eg = new ExcelGenerator("new_output.xls");
//        er.readSheet(0);
//        int[] indices = new int[]{1, 2, 3};
//        String[][] content = er.getSheet();
//        ArrayList<ArrayList<String>> arrCon = spliteCells(content, indices, ",");
//        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
//        tableArr.add(toTable(arrCon, content, indices));
//        eg.createFile(tableArr);
        CSVOrganizer organizer = new CSVOrganizer("input.csv");
        organizer.process(1);
        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
        tableArr.add(organizer.exportAsExcelTable());
        eg.createFile(tableArr);
    }
}
