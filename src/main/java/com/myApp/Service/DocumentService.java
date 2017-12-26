package com.myApp.Service;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Service;

@Service
public class DocumentService {
        HashMap<String, String> replacements = new HashMap<String, String>();

    public HashMap<String, String> getReplacements() {
        replacements.put("${varDate}", "10.12.2017");
        replacements.put("${varStatementNumber}", "1/102");
        replacements.put("${varFullName}", "Diāna Samarska");
        replacements.put("${varHiredDate}", "01.01.2017");
        replacements.put("${varProfession}", "programmētāju");
        replacements.put("${varProfessionCode}", "25001");
        replacements.put("${varSalaryDigits}", "1000");
        replacements.put("${varSalaryString}", "viens tūkstotis");
        replacements.put("${varSigningPerson}", "Zanda Arnava");
        replacements.put("${varSigningPersonsPosition}", "HR vadītāja");
        return replacements;
    }
  public String create(String name, String surname){
    DocumentService doc = new DocumentService();
    doc.getReplacements();

    doc.searchAndReplace("/home/bi/Downloads/Sample.doc", "/home/bi/Downloads/"+name+surname+".doc", doc.replacements);
return "Done";
}

public void searchAndReplace(String inputFilename, String outputFilename, HashMap<String, String> map) {

       File inputFile = null;
        File outputFile = null;
        FileInputStream fileIStream = null;
        FileOutputStream fileOStream = null;
        BufferedInputStream bufIStream = null;
        BufferedOutputStream bufOStream = null;
        POIFSFileSystem fileSystem = null;
        HWPFDocument document = null;
        Range docRange = null;
        Paragraph paragraph = null;
        CharacterRun charRun = null;
        Set<String> keySet = null;
        Iterator<String> keySetIterator = null;
        int numParagraphs = 0;
        int numCharRuns = 0;
        String text = null;
        String key = null;
        String value = null;

        try {
            inputFile = new File(inputFilename);
            fileIStream = new FileInputStream(inputFile);
            bufIStream = new BufferedInputStream(fileIStream);
            fileSystem = new POIFSFileSystem(bufIStream);

            document = new HWPFDocument(fileSystem);

            docRange = document.getRange();

            numParagraphs = docRange.numParagraphs();

            keySet =replacements.keySet();

            for(int i = 0; i < numParagraphs; i++) {
                paragraph = docRange.getParagraph(i);
                text = paragraph.text();
                numCharRuns = paragraph.numCharacterRuns();

                for(int j = 0; j < numCharRuns; j++) {
                    charRun = paragraph.getCharacterRun(j);
                    text = charRun.text();
                    System.out.println("Character Run text: " + text);
                    keySetIterator = keySet.iterator();

                    while(keySetIterator.hasNext()) {
                        key = keySetIterator.next();

                        if(text.contains(key)) {
                            value = replacements.get(key);
                            charRun.replaceText(key, value);
                            docRange = document.getRange();
                            paragraph = docRange.getParagraph(i);
                            charRun = paragraph.getCharacterRun(j);
                            text = charRun.text();
                        }
                    }
                }
            }

            bufIStream.close();
            bufIStream = null;

            outputFile = new File(outputFilename);
            fileOStream = new FileOutputStream(outputFile);
            bufOStream = new BufferedOutputStream(fileOStream);

            document.write(bufOStream);

        }
        catch(Exception ex) {
            System.out.println("Caught an: " + ex.getClass().getName());
            System.out.println("Message: " + ex.getMessage());
            System.out.println("Stacktrace follows.............");
            ex.printStackTrace(System.out);
        }
        finally {
            if (bufIStream != null) {
                try {
                    bufIStream.close();
                    bufIStream = null;
                } catch (Exception ex) {
                    // I G N O R E //
                }
            }
            if (bufOStream != null) {
                try {
                    bufOStream.flush();
                    bufOStream.close();
                    bufOStream = null;
                } catch (Exception ex) {
                    // I G N O R E //
                }
            }
            System.out.println("New Document is ready");
        }
}

}