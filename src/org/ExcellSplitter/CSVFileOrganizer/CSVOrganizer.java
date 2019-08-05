package org.ExcellSplitter.CSVFileOrganizer;

import org.ExcellSplitter.ExcelProcessor.ExcelTable;

import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.StringTokenizer;

public class CSVOrganizer {
    private File file;
    private ArrayList<String> tags;
    private ArrayList<ArrayList<String>> content;
    private Hashtable<String, Integer> tagsIndices;

    public CSVOrganizer (String address){
        if (!address.endsWith(".csv")) {
            throw new IllegalArgumentException();
        }
        file = new File(address);
    }

    public void process (int leftColumns) throws IOException {
        content = new ArrayList<>();
        tags = new ArrayList<>();
        tagsIndices = new Hashtable<>();

        FileInputStream fis = new FileInputStream(file);
        InputStreamReader isr = new InputStreamReader(fis);
        BufferedReader bufferedReader = new BufferedReader(isr);

        String firstLine = bufferedReader.readLine();
        StringTokenizer stk = new StringTokenizer(firstLine, ",");
        int oldTagNum = 0;
        while (stk.hasMoreTokens() && oldTagNum < leftColumns) {
            String oldTag = stk.nextToken();
            tags.add(oldTag);
            tagsIndices.put(oldTag, oldTagNum);
            oldTagNum++;
        }

        String line = bufferedReader.readLine();
        while (line != null && line.length() > 1) {
            line = line.replace("\"", "");
            StringTokenizer lineToken = new StringTokenizer(line, ",");
            ArrayList<String> row = new ArrayList<>();
            for (int i = 0; i < leftColumns; i++) {
                row.add(lineToken.nextToken());
            }

            HashSet<String> rowContent = new HashSet<>();
            while (lineToken.hasMoreTokens()) {
                String ele = lineToken.nextToken();
                if (ele.length() < 1) {
                    continue;
                }
                rowContent.add(ele);
                if (!tagsIndices.containsKey(ele)) {
                    tags.add(ele);
                    tagsIndices.put(ele, oldTagNum);
                    oldTagNum++;
                }
            }

            for (int i = leftColumns; i < tags.size(); i++) {
                if (rowContent.contains(tags.get(i))) {
                    row.add(tags.get(i));
                } else {
                    row.add("");
                }
            }
            content.add(row);
            line = bufferedReader.readLine();
        }

        for (int i = 0; i < content.size(); i++) {
            if (content.get(i).size() < tags.size()) {
                while (content.get(i).size() < tags.size()) {
                    content.get(i).add("");
                }
            }
        }
    }

    public ExcelTable exportAsExcelTable () {
        ExcelTable table = new ExcelTable(tags);
        table.addRows(content.size());
        for (int i = 0; i < content.size(); i++) {
            for (int j = 0; j < content.get(i).size(); j++) {
                table.addCell(i, j, content.get(i).get(j));
            }
        }
        return table;
    }
}
