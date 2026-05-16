package com.geno;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;

public class DocumentGenerator {

    public static void main(String[] args) throws Exception {
        // 1. Lire le JSON
        ObjectMapper mapper = new ObjectMapper();
        JsonNode root = mapper.readTree(new File("projet_geno/src/main/resources/document.json"));

        // 2. Créer le document
        XWPFDocument doc = new XWPFDocument();

        // Titre principal
        String title = root.get("title").asText();
        XWPFParagraph titlePara = doc.createParagraph();
        titlePara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = titlePara.createRun();
        titleRun.setText(title);
        titleRun.setBold(true);
        titleRun.setFontSize(24);
        titleRun.setFontFamily("Arial");

        // En-tête avec le nom de l'auteur
        String author = root.get("author").asText();
        XWPFHeader header = doc.createHeader(org.apache.poi.wp.usermodel.HeaderFooterType.DEFAULT);
        XWPFParagraph headerParagraph = header.createParagraph();
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun headerRun = headerParagraph.createRun();
        headerRun.setText("By: " + author);
        headerRun.setFontFamily("Arial");
        headerRun.setFontSize(12);

        // Pied de page avec les numéros de page
        XWPFFooter footer = doc.createFooter(org.apache.poi.wp.usermodel.HeaderFooterType.DEFAULT);
        XWPFParagraph footerParagraph = footer.createParagraph();
        footerParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun footerRun = footerParagraph.createRun();
        footerRun.setText("Page ");
        footerParagraph.getCTP().addNewFldSimple().setInstr("PAGE \\* MERGEFORMAT");
        footerRun = footerParagraph.createRun();
        footerRun.setText(" of ");
        footerParagraph.getCTP().addNewFldSimple().setInstr("NUMPAGES \\* MERGEFORMAT");

        // 3. Parcourir les sections
        for (JsonNode section : root.get("sections")) {
            String heading = section.get("heading").asText();
            int level = section.get("level").asInt();
            String content = section.get("content").asText();

            // Heading avec style selon le niveau
            XWPFParagraph headingPara = doc.createParagraph();
            headingPara.setStyle("Heading" + level);
            XWPFRun headingRun = headingPara.createRun();
            headingRun.setText(heading);

            // Contenu
            XWPFParagraph contentPara = doc.createParagraph();
            XWPFRun contentRun = contentPara.createRun();
            contentRun.setText(content);

            // Styles optionnels depuis le JSON
            if (section.has("bold")) {
                contentRun.setBold(section.get("bold").asBoolean());
            }
            if (section.has("fontSize")) {
                contentRun.setFontSize(section.get("fontSize").asInt());
            }
            if (section.has("color")) {
                contentRun.setColor(section.get("color").asText());
            }
        }

        // 4. Sauvegarder
        try (FileOutputStream out = new FileOutputStream("output.docx")) {
            doc.write(out);
        }
        doc.close();
        System.out.println("Document généré : output.docx");
    }
}