//package org.example;
//
//import org.apache.poi.xwpf.usermodel.*;
//import org.apache.poi.util.Units;
//import javax.imageio.ImageIO;
//import java.awt.image.*;
//import java.awt.color.ColorSpace;
//import java.io.*;
//
//public class AdjustImageInDocx {
//
//    public static void main(String[] args) throws Exception {
//        String inputDocx  = "input.docx";
//        String outputDocx = "output.docx";
//
//        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(inputDocx));
//             FileOutputStream out = new FileOutputStream(outputDocx)) {
//
//            // Duyệt qua tất cả paragraph → run → picture
//            for (XWPFParagraph p : doc.getParagraphs()) {
//                for (XWPFRun r : p.getRuns()) {
//                    for (XWPFPicture pic : r.getEmbeddedPictures()) {
//                        XWPFPictureData picData = pic.getPictureData();
//                        if (picData == null) continue;
//
//                        // Đọc ảnh gốc
//                        BufferedImage img = ImageIO.read(picData.getInputStream());
//
//                        // Tính % tăng → 0% = giữ nguyên, 40% = tăng mạnh
//                        float percent = 0.40f; // bạn có thể truyền tham số từ 0 đến 0.4
//
//                        // Tăng brightness và contrast
//                        BufferedImage adjusted = adjustBrightnessContrast(img, percent, percent);
//
//                        // Ghi đè picture data mới (giữ nguyên format jpg/png)
//                        ByteArrayOutputStream baos = new ByteArrayOutputStream();
//                        ImageIO.write(adjusted, picData.getType() == XWPFPictureData.PNG ? "png" : "jpg", baos);
//
//                        // Thay thế dữ liệu ảnh cũ
//                        picData.setData(baos.toByteArray());
//                    }
//                }
//            }
//
//            doc.write(out);
//            System.out.println("Đã xử lý xong → " + outputDocx);
//        }
//    }
//
//    /**
//     * Tăng brightness và contrast theo phần trăm (0.0 → 0.4 tương ứng 0-40%)
//     * Đây là cách đơn giản, hiệu ứng gần giống Word
//     */
//    public static BufferedImage adjustBrightnessContrast(BufferedImage src, float brightnessPercent, float contrastPercent) {
//        // brightnessPercent: 0.0 = không đổi → 0.4 ≈ +40%
//        // contrastPercent:  tương tự
//
//        float brightness = 1.0f + brightnessPercent * 1.0f; // tăng sáng
//        float contrast   = 1.0f + contrastPercent * 2.0f;   // tăng contrast mạnh hơn
//
//        // Công thức RescaleOp: output = (input - 0.5) * contrast + 0.5 + brightnessOffset
//        // Nhưng ta đơn giản hóa
//
//        RescaleOp op = new RescaleOp(contrast,  // scale factor (contrast)
//                (brightness - 1.0f) * 128,     // offset ≈ brightness
//                null);
//
//        return op.filter(src, null);
//    }
//}