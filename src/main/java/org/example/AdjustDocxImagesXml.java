package org.example;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLuminanceEffect;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class AdjustDocxImagesXml {

    private static final String INPUT_DIR  = "input";
    private static final String OUTPUT_DIR = "output";

    private static final int INC_BRIGHT = 40 * 1000;
    private static final int INC_CONTRAST = 40 * 1000;

    public static void main(String[] args) throws Exception {

        File inDir = new File(INPUT_DIR);
        File outDir = new File(OUTPUT_DIR);

        if (!outDir.exists()) outDir.mkdirs();

        File[] files = inDir.listFiles((d, n) -> n.toLowerCase().endsWith(".docx"));
        if (files == null) return;

        for (File in : files) {
            File out = new File(outDir, in.getName());
            processDocx(in, out);
        }
    }

    private static void processDocx(File input, File output) throws Exception {

        try (FileInputStream fis = new FileInputStream(input);
             XWPFDocument doc = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(output)) {

            processBodyElements(doc.getBodyElements());

            for (XWPFHeader h : doc.getHeaderList()) {
                processBodyElements(h.getBodyElements());
            }

            for (XWPFFooter f : doc.getFooterList()) {
                processBodyElements(f.getBodyElements());
            }

            doc.write(fos);
        }
    }

    private static void processBodyElements(List<IBodyElement> elements) {

        for (IBodyElement el : elements) {

            if (el instanceof XWPFParagraph) {
                processParagraph((XWPFParagraph) el);

            } else if (el instanceof XWPFTable) {
                processTable((XWPFTable) el);
            }
        }
    }

    private static void processParagraph(XWPFParagraph p) {

        for (XWPFRun r : p.getRuns()) {
            for (XWPFPicture pic : r.getEmbeddedPictures()) {
                adjustPicture(pic);
            }
        }
    }

    private static void processTable(XWPFTable table) {

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                processBodyElements(cell.getBodyElements());
            }
        }
    }

    private static void adjustPicture(XWPFPicture picture) {

        CTPicture ctPic = picture.getCTPicture();
        if (ctPic == null) return;

        CTBlipFillProperties blipFill = ctPic.getBlipFill();
        if (blipFill == null) blipFill = ctPic.addNewBlipFill();

        CTBlip blip = blipFill.getBlip();
        if (blip == null) blip = blipFill.addNewBlip();

        List<CTLuminanceEffect> lumList = blip.getLumList();

        if (!lumList.isEmpty()) {
            CTLuminanceEffect lum = lumList.get(0);

            int b = lum.isSetBright() ? (int) lum.getBright() : 0;
            int c = lum.isSetContrast() ? (int) lum.getContrast() : 0;

            lum.setBright(clamp(b + INC_BRIGHT));
            lum.setContrast(clamp(c + INC_CONTRAST));

        } else {
            CTLuminanceEffect lum = blip.addNewLum();
            lum.setBright(clamp(INC_BRIGHT));
            lum.setContrast(clamp(INC_CONTRAST));
        }
    }

    private static int clamp(int v) {
        if (v > 100000) return 100000;
        if (v < -100000) return -100000;
        return v;
    }
}
