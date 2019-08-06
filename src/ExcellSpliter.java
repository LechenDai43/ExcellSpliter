import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Pane;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import jxl.write.WriteException;
import org.ExcellSplitter.CSVFileOrganizer.CSVOrganizer;
import org.ExcellSplitter.ExcelProcessor.CellSplitor;
import org.ExcellSplitter.ExcelProcessor.ExcelGenerator;
import org.ExcellSplitter.ExcelProcessor.ExcelReader;
import org.ExcellSplitter.ExcelProcessor.ExcelTable;
import javafx.scene.control.Button;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.StringTokenizer;

public class ExcellSpliter extends Application {
    private File file;
    private Button bt1 = new Button("Choose Input File");
    private Button bt2 = new Button("Start Splitting");
    private Label label = new Label(),
            address = new Label("Your file name will show up here.");
    private Label colL = new Label("Enter columns: ");
    private TextField colT = new TextField();
    private Label delL = new Label("Enter delimiter: ");
    private TextField delT = new TextField();
    private FileChooser fileChooser;

    public static void main (String[] args) throws IOException, WriteException {
//        ExcelReader er = new ExcelReader("input.xls");
//        ExcelGenerator eg = new ExcelGenerator("new_output.xls");
//        er.readSheet(0);
//        int[] indices = new int[]{1, 2, 3};
//        String[][] content = er.getSheet();
//        ArrayList<ArrayList<String>> arrCon = spliteCells(content, indices, ",");
//        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
//        tableArr.add(toTable(arrCon, content, indices));
//        eg.createFile(tableArr);
//        CSVOrganizer organizer = new CSVOrganizer("input.csv");
//        organizer.process(1);
//        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
//        tableArr.add(organizer.exportAsExcelTable());
//        eg.createFile(tableArr);
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        Scene scene = new Scene(getPane(), 450, 350);
        primaryStage.setTitle("Excel Cell Split");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private Pane getPane() {
        VBox result = new VBox(10);
        result.setAlignment(Pos.CENTER);
        HBox hBox1 = new HBox(5),
                hBox2 = new HBox(5);

        result.getChildren().addAll(bt1, address, hBox1, hBox2, bt2, label);


        hBox1.getChildren().addAll(colL, colT);
        hBox1.setAlignment(Pos.CENTER);
        hBox1.setMargin(colL, new Insets(3));
        hBox1.setMargin(colT, new Insets(3));


        hBox2.getChildren().addAll(delL, delT);
        hBox2.setAlignment(Pos.CENTER);
        hBox2.setMargin(delL, new Insets(3));
        hBox2.setMargin(delT, new Insets(3));

        fileChooser = new FileChooser();

        bt1.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                file = fileChooser.showOpenDialog(null);
                address.setText(file.getName());
                if (file.getName().endsWith(".xls")) {
                    colL.setText("Enter target columns (separate by ,): ");
                    delT.setEditable(true);
                } else if (file.getName().endsWith(".csv")){
                    colL.setText("Enter how many columns to ignore: ");
                    delT.setEditable(false);
                } else {
                    colL.setText("Enter columns: ");
                    delT.setEditable(true);
                    address.setText("This file's format is not valid.");
                }
            }
        });

        bt2.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                String addIn = file.getAbsolutePath();
                String addOut = addIn.substring(0, addIn.lastIndexOf("."));
                addOut += "_split.xls";
                if (addIn.endsWith(".xls")) {
                    try {
                        forExcel(addOut, colT.getText(), delT.getText());
                    } catch (IOException e) {
                        e.printStackTrace();
                    } catch (WriteException e) {
                        e.printStackTrace();
                    }
                } else {
                    try {
                        forCSV(addOut, colT.getText());
                    } catch (IOException e) {
                        e.printStackTrace();
                    } catch (WriteException e) {
                        e.printStackTrace();
                    }
                }
                label.setText("Your output is at: " + addOut);
            }
        });

        for (int i = 0; i < result.getChildren().size(); i++) {
            Node child = result.getChildren().get(i);
            result.setMargin(child, new Insets(5));
        }
        return result;
    }

    private void forExcel (String out, String sIndices, String delimiter) throws IOException, WriteException {
        ExcelReader er = new ExcelReader(file.getAbsoluteFile());
        ExcelGenerator eg = new ExcelGenerator(out);
        er.readSheet(0);
        int[] indices;
        StringTokenizer stk = new StringTokenizer(sIndices, ",");
        indices = new int[stk.countTokens()];
        int i = 0;
        while (stk.hasMoreTokens()) {
            int j = Integer.parseInt(stk.nextToken()) - 1;
            indices[i] = j;
            i++;
        }
        String[][] content = er.getSheet();
        ArrayList<ArrayList<String>> arrCon = CellSplitor.spliteCells(content, indices, delimiter);
        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
        tableArr.add(CellSplitor.toTable(arrCon, content, indices));
        eg.createFile(tableArr);
    }

    private void forCSV (String out, String toNum) throws IOException, WriteException {
        ExcelGenerator eg = new ExcelGenerator(out);
        CSVOrganizer organizer = new CSVOrganizer(file.getAbsolutePath());
        organizer.process(Integer.parseInt(toNum));
        ArrayList<ExcelTable> tableArr = new ArrayList<ExcelTable>();
        tableArr.add(organizer.exportAsExcelTable());
        eg.createFile(tableArr);
    }
}
