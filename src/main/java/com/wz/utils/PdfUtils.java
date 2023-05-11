package com.wz.utils;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author 王哲
 * @ClassName PdfUtils
 * @create 2023--五月--下午10:51
 * @Description PDF工具类
 * @Version V1.0
 */
public class PdfUtils {

    /**
     * pdf转成word
     * @param pdfFile
     * @param response
     * @throws IOException
     */
    public static void pdfToWord(MultipartFile pdfFile, HttpServletResponse response) throws IOException {

        // 读取PDF文件内容
        PDDocument pdf = PDDocument.load(pdfFile.getInputStream());
        PDFTextStripper stripper = new PDFTextStripper();
        String text = stripper.getText(pdf);
        pdf.close();

        // 将PDF内容写入Word文档
        XWPFDocument document = new XWPFDocument();
        String[] textPerPage = text.split("\u000C"); // 根据换页符拆分文本
        for (int i = 0; i < textPerPage.length; i++) {
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText(textPerPage[i]);
            if (i < textPerPage.length - 1) {
                // 添加分页符
                run.addBreak();
                run.addBreak();
            }
        }

        // 输出Word文件
        String fileName = pdfFile.getOriginalFilename().replace(".pdf", ".docx");
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", fileName);
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
        OutputStream out = response.getOutputStream();
        document.write(out);
        out.close();
        document.close();
    }
}
