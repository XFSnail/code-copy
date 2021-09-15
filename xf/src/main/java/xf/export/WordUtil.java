package xf.export;

import com.lowagie.text.Cell;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import java.awt.Color;
import java.io.File;
import org.springframework.core.io.ClassPathResource;

/**
 * word导出工具
 * 通过 itext 实现
 *
 * @author xiong fan
 * @create 2021-09-03 14:43
 */
public class WordUtil {

    public static final String font1 = "static/simhei1.ttf";
    public static final String font2 = "static/simsong.ttf";

    public static Document createDocumet() throws Exception {
        Document document = new Document();
        document.setPageSize(PageSize.A4);
        return document;
    }

    public static Font getFont(String fontPath) throws Exception {
        ClassPathResource classPathResource = new ClassPathResource(fontPath);
        String path = classPathResource.getURL().getPath();
        return new Font(BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED));
    }

    public static Font getFont(String fontPath, Integer size, Integer style) throws Exception {
        ClassPathResource classPathResource = new ClassPathResource(fontPath);
        String path = classPathResource.getURL().getPath();
        return new Font(BaseFont.createFont(path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED), size, style);
    }

    public static StringBuffer wordFilePath() {
        // 设置pdf生成的路径
        StringBuffer str = new StringBuffer(System.getProperty("java.io.tmpdir"));
        // linux路径
        if (System.getProperty("os.name").toLowerCase().indexOf("linux") > -1) {
            String dir = "/opt/word/";
            str = new StringBuffer(dir);
            File dirFile = new File(dir);
            if (!dirFile.exists()) {
                dirFile.mkdirs();
            }
        }
        return str;
    }

    //段落（标题）居中
    public static void paragrahCenter(Document document, String content, Font font) throws Exception {
        Paragraph paragraph = new Paragraph(content, font);
        paragraph.setAlignment(Element.ALIGN_CENTER);
        document.add(paragraph);
    }

    public static Table craeteWordTable(int numColumns, float totalWidth, float padding,
        int... widths) throws DocumentException {
        Table table = new Table(numColumns);
        table.setWidths(widths);//设置系列所占比例
        table.setWidth(totalWidth);
        table.setAlignment(Element.ALIGN_CENTER);//居中显示
        table.setAlignment(Element.ALIGN_MIDDLE);//垂直居中显示
        table.setPadding(padding);
        return table;
    }

    public static void addCell(Table table, String content, int colspan, int rowspan, Color backgroundColor, Font font)
        throws Exception {
        Cell cell = new Cell(new Phrase(content, font));
        cell.setColspan(colspan);//设置当前单元格占据的列数
        cell.setRowspan(rowspan);
        cell.setVerticalAlignment(Element.ALIGN_CENTER);
        cell.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell.setBackgroundColor(backgroundColor);
        table.addCell(cell);
    }

    /*public static PdfPTable craetePdfPTable(int numColumns, float totalWidth, float spacing, int... widths) throws DocumentException {
        PdfPTable table = new PdfPTable(numColumns);
        table.setWidths(widths);
        table.setTotalWidth(totalWidth);
        table.setSpacingBefore(spacing);
        table.setLockedWidth(true);
        return table;
    }

    public static void addCell(PdfPTable table, com.itextpdf.text.Font font, String content, int horizontalAlignment, boolean isHideBorder, int colspan, float paddingLeft, float minHeight) {
        PdfPCell cell = new PdfPCell(new Paragraph(content, font));
        cell.setPaddingLeft(paddingLeft);
        cell.setMinimumHeight(minHeight);
        //设置单元格的垂直对齐方式
        cell.setVerticalAlignment(PdfPCell.ALIGN_MIDDLE);
        //设置单元格的水平对齐方式
        cell.setHorizontalAlignment(horizontalAlignment);
        if (colspan > 0) {
            //合并单元格
            cell.setColspan(colspan);
        }
        if (isHideBorder) {
            //隐藏单元格边框
            cell.disableBorderSide(15);
        }
        table.addCell(cell);
    }*/
}