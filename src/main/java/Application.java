import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class Application {
    public static void main(String[] args) throws InvalidFormatException, IOException {
        System.out.println("Start...");

        String inputFilePath = "file-sample_500kB_dv.docx";
        String outputFilePath = "file-sample_500kB_dv_result.docx";
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(inputFilePath));
        replaceText(doc, "${text}", "wxyz");
        doc.write(new FileOutputStream(outputFilePath));

        System.out.println("Finish...");
    }

    private static void replaceText(XWPFDocument doc, String originalText, String replaceText) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(originalText)) {
                        text = text.replace(originalText, replaceText);
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : doc.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains(originalText)) {
                                text = text.replace(originalText, replaceText);
                                r.setText(text,0);
                            }
                        }
                    }
                }
            }
        }
    }

}
