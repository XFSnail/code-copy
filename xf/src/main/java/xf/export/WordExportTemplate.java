package xf.export;

import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Table;
import com.lowagie.text.rtf.RtfWriter2;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import javax.servlet.http.HttpServletResponse;
import org.springframework.stereotype.Component;

/**
 * 房屋建设导出模板
 *
 * @author xiong fan
 * @create 2021-08-27 13:53
 */
@Component
public class WordExportTemplate {

    // 黑体
    private static Font fontHei;
    // 宋体
    private static Font fontSong;
    // 宋体
    private static Font fontSongBold;

    /**
     * 农村房屋导出模板
     *
     * @param map 农村房屋数据
     * @return
     * @Author xiong fan
     * @Date 15:53 2021/8/27
     */
    public static void exportCountryHouse(Map<String, Object> data, HttpServletResponse response) throws Exception {
        Map<String, Object> map = data != null ? (Map) data.get("real") : new HashMap<>();
        Document document = WordUtil.createDocumet();

        // 设置为浏览器的下载位置
        response.setContentType("application/msword;charset=UTF-8");
        response.setCharacterEncoding("gbk");
        String fileName = new String(("example" + System.currentTimeMillis() + ".doc")
            .getBytes("GBK"), "iso8859-1");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
        RtfWriter2.getInstance(document, response.getOutputStream());
        // 下载至本地
        /*String path = "E:/" + System.currentTimeMillis() + ".docx";
        RtfWriter2.getInstance(document, new FileOutputStream(path));*/

        fontHei = WordUtil.getFont(WordUtil.font1, 14, Font.NORMAL);
        fontSong = WordUtil.getFont(WordUtil.font2, 9, Font.NORMAL);
        fontSongBold = WordUtil.getFont(WordUtil.font2, 9, Font.BOLD);

        try {
            document.open();
            // 塞数据
            WordUtil.paragrahCenter(document, "title", fontHei);
            document.add(stuffData(map));
        } catch (IOException e) {
            throw new Exception(e);
        } finally {
            document.close();
        }
    }


    private static Table stuffData(Map<String, Object> map) throws Exception {
        // 建筑层数
        String cs = map.get("cs") == null ? "" : map.get("cs").toString();
        // 建筑面积
        String dcmj = map.get("dcmj") == null ? "" : map.get("dcmj").toString();
        // 建筑年代 1、1980年及以前；2、1981-1990年；3、1991-2000年；4、2001-2010年；5、2011-2015年；6、2016年及以后
        String jzndQgStr = map.get("jzndQg") == null ? "" : map.get("jzndQg").toString();
        String jzndQg = "";
        switch (jzndQgStr) {
            case "1":
                jzndQg = "1980年及以前";
                break;
            case "2":
                jzndQg = "1981-1990年";
                break;
            case "3":
                jzndQg = "1991-2000年";
                break;
            case "4":
                jzndQg = "2001-2010年";
                break;
            case "5":
                jzndQg = "2011-2015年";
                break;
            case "6":
                jzndQg = "2016年及以后";
                break;
            default:
        }

        Table table = WordUtil.craeteWordTable(12, 100, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1);
        table.setTop(-200);
        table.setOffset(0);
        // 未选中的复选框：□
        // 选中的复选框：☑
        WordUtil.addCell(table, "未选中的复选框：□  ", 12, 1, null, fontSongBold);
        WordUtil.addCell(table, "选中的复选框：☑  多来几个：☑ □复选框：□  ", 12, 2, null, fontSongBold);

        WordUtil.addCell(table, "test第二部分：建筑信息", 12, 2, null, fontSongBold);
        WordUtil.addCell(table, "2.1 建筑层数", 2, 1, null, fontSong);
        WordUtil.addCell(table, cs + "层", 4, 1, null, fontSong);
        WordUtil.addCell(table, "2.2 建筑面积", 2, 1, null, fontSong);
        WordUtil.addCell(table, dcmj + "平方米", 4, 1, null, fontSong);
        WordUtil.addCell(table, "2.3 建造年代", 2, 1, null, fontSong);
        WordUtil.addCell(table, jzndQg, 10, 1, null, fontSong);
        WordUtil.addCell(table, "2.4 结构类型", 2, 1, null, fontSong);
        return table;
    }
}