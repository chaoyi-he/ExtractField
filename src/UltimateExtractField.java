import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;

/**
 * Created by hechaoyi on 16/2/22.
 */
public class UltimateExtractField {

    String inputFileString;
    String outFileString;
    String flag;

    File inputFile;
    FileInputStream fileInputStream;

    File outputFile;
    FileWriter fileWriter;
    BufferedWriter bufferWriter;

    Range range;
    POIFSFileSystem poifsFileSystem;
    HWPFDocument hw;

    public UltimateExtractField(String inputFileString, String outFileString, String flag) {
        this.inputFileString = inputFileString;
        this.outFileString = outFileString;
        this.flag = flag;
    }

    /**
     * 读取doc文件表格
     * @param range
     * @param fileout
     * @throws Exception
     */
    private void readTableInfo(Range range, BufferedWriter fileout) throws Exception {
        TableIterator ti = new TableIterator(range);
        while (ti.hasNext()) {
            Table table = ti.next();
            int numberRows = table.numRows();
            for (int i = 0; i < numberRows; i++) {
                TableRow tableRow = table.getRow(i);
                int numCell = tableRow.numCells();

                if (flag.equals("1")) {
                    if (numCell <= 1) {
                        continue;
                    } else {
                        String str = tableRow.getCell(2).text().trim() + "*" + tableRow.getCell(0).text().trim() + "\n";
                        System.out.println(str);
                        fileout.write(str);
                    }
                } else if (flag.equals("2")) {
                    if (numCell == 2) {
                        String str = tableRow.getCell(0).text().trim() + "*" + tableRow.getCell(1).text().trim() + "\n";
                        fileout.write(str);
                    }
                }
            }
        }
    }

    /**
     *
     * 获取输出文件流
     */
    private BufferedWriter getFileBufferWriter() {


        try {
            outputFile = new File(outFileString);
            if (outputFile.exists()) {
                outputFile.delete();
                outputFile.createNewFile();
            }
            fileWriter = new FileWriter(outputFile, true);
            bufferWriter = new BufferedWriter(fileWriter);

        } catch (IOException e) {
            e.printStackTrace();
        }

        return bufferWriter;
    }

    /**
     * 得到poi文件块
     * @return
     */
    private Range getRange() {
        inputFile = new File(inputFileString);
        try {
            fileInputStream = new FileInputStream(inputFile);
            poifsFileSystem = new POIFSFileSystem(fileInputStream);
            hw = new HWPFDocument(poifsFileSystem);
            range = hw.getRange();

        } catch (IOException e) {
            e.printStackTrace();
        }

        return range;
    }


    private void close() {
        try {
            bufferWriter.close();
            fileWriter.close();
            fileInputStream.close();
            poifsFileSystem.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void main(String args[]) {

        UltimateExtractField ult = new UltimateExtractField(args[0], args[1], "1");
        UltimateExtractField ultOther = new UltimateExtractField(args[2], args[3], "2");

        try {
            ult.readTableInfo(ult.getRange(), ult.getFileBufferWriter());
            ultOther.readTableInfo(ultOther.getRange(), ultOther.getFileBufferWriter());
        } catch (Exception e) {
            e.printStackTrace();
        }
        ult.close();
        ultOther.close();


    }

}
