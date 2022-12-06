package com.vioplugins.pluginlabreport;

import com.intellij.openapi.actionSystem.AnAction;
import com.intellij.openapi.actionSystem.AnActionEvent;
import com.intellij.openapi.actionSystem.PlatformDataKeys;
import com.intellij.openapi.editor.Editor;
import com.intellij.openapi.ui.Messages;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


import org.jetbrains.annotations.NotNull;

import javax.swing.text.Document;
import java.io.FileOutputStream;
import java.io.IOException;


public class CreaterReportPlugin extends AnAction {
    @Override
    public void actionPerformed(@NotNull AnActionEvent e) {
        Editor highlighting = e.getData(PlatformDataKeys.EDITOR);
        assert highlighting != null;
        String highlightStr = highlighting.getSelectionModel().getSelectedText();
        try {
            createDocxFile(highlightStr);
        } catch (Exception exception) {
            Messages.showMessageDialog("Error with docx file", "Error Message", Messages.getInformationIcon());
        }


    }

    @Override
    public boolean isDumbAware() {
        return false;
    }

    public void createDocxFile(String highlightStr) throws IOException {
        String fileName = "D:\\devtools5lab\\pluginLabReport\\report.docx";


        try (XWPFDocument doc = new XWPFDocument()) {

            // create a paragraph
            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.CENTER);

            // set font
            XWPFRun r1 = p1.createRun();
            r1.setBold(true);
            r1.setFontSize(22);
            r1.setFontFamily("New Roman");
            r1.setText("Lab 5.");
            r1.addBreak();
            r1.setText("Report for IT-project.");
            XWPFParagraph p2 = doc.createParagraph();
            XWPFRun r2 = p2.createRun();
            p2.setAlignment(ParagraphAlignment.LEFT);
            r2.addBreak();
            r2.setFontSize(14);
            r2.setBold(false);
            r2.setText(highlightStr);

            // save it to .docx file
            try (FileOutputStream out = new FileOutputStream(fileName)) {
                doc.write(out);

            }

        }
    }
}


