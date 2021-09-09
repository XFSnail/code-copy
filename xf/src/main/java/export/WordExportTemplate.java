package export;

import com.lowagie.text.Document;
import com.lowagie.text.Font;
import com.lowagie.text.Table;
import com.lowagie.text.rtf.RtfWriter2;
import java.util.HashMap;
import java.util.Map;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.lang.StringUtils;
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
        response.setCharacterEncoding("utf-8");
        response.setContentType("application/msword");
        String fileName = new String("农村房屋调查信息.doc".getBytes("GBK"), "iso8859-1");
        response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
        RtfWriter2.getInstance(document, response.getOutputStream());

        fontHei = WordUtil.getFont(WordUtil.font1);
        fontSong = WordUtil.getFont(WordUtil.font2);
        fontHei.setSize(14);
        fontSong.setSize(9);
        fontSongBold = WordUtil.getFont(WordUtil.font2, 9, Font.BOLD);

        String houseType = map.get("houseType") == null ? "" : map.get("houseType").toString();
        // 农村房屋类别： 1、住宅-独立住宅；2、住宅-集合住宅；3、非住宅
        int type = 1;
        if ("1".equals(houseType)) {
            type = map.get("fwlx") == null ? 1 : Integer.parseInt(map.get("fwlx").toString());
        } else if ("2".equals(houseType)) {
            type = 3;
        }
        String title = "农村住宅建筑调查信息采集表";
        if (type == 1) {
            title += "（独立住宅）";
        } else if (type == 2) {
            title += "（集合住宅）";
        }

        // 第二部分
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

        // 结构类型 1、砖石结构；2、土木结构；3、混杂结构；4、窑洞；5、钢筋混凝土结构；6、钢结构；99、其他
        String jglxQgStr = map.get("jglxQg") == null ? "" : map.get("jglxQg").toString();
        String jglxQg = "";
        switch (jglxQgStr) {
            case "1":
                jglxQg = "砖石结构";
                break;
            case "2":
                jglxQg = "土木结构";
                break;
            case "3":
                jglxQg = "混杂结构";
                break;
            case "4":
                jglxQg = "窑洞";
                break;
            case "5":
                jglxQg = "钢筋混凝土结构";
                break;
            case "6":
                jglxQg = "钢结构";
                break;
            case "99":
                jglxQg = map.get("qtjglxQg") == null ? "" : map.get("qtjglxQg").toString();
                break;
            default:
        }
        // 建造方式 1、自行建造；2、建筑工匠建造；3、有资质的施工队伍建造；99、其他
        String jzfsQgStr = map.get("jzfsQg") == null ? "" : map.get("jzfsQg").toString();
        String jzfsQg = "";
        switch (jzfsQgStr) {
            case "1":
                jzfsQg = "自行建造";
                break;
            case "2":
                jzfsQg = "建筑工匠建造";
                break;
            case "3":
                jzfsQg = "有资质的施工队伍建造";
                break;
            case "99":
                jzfsQg = map.get("qtjzfsQg") == null ? "" : map.get("qtjzfsQg").toString();
                break;
            default:
        }

        // 建造用途-教育设施：是否为"中小学幼儿园教学用房及学生宿舍、食堂"
        String jyssQgStr = map.get("jyssQg") == null ? "" : map.get("jyssQg").toString();
        String jyssQg = "1".equals(jyssQgStr) ? "是" : "0".equals(jyssQgStr) ? "否" : "";
        // 建造用途-医疗卫生：是否为"具有外科手术室或急诊科的乡镇卫生院医疗用房"
        String ylssQgStr = map.get("ylssQg") == null ? "" : map.get("ylssQg").toString();
        String ylssQg = "1".equals(ylssQgStr) ? "是" : "0".equals(ylssQgStr) ? "否" : "";

        // 建造用途 1、教育设施；2、医疗卫生；3、行政办公；4、文化设施；5、养老服务；6、批发零售；7、餐饮服务；
        // 8、住宿宾馆；9、休闲娱乐；10、宗教场所；11、农贸市场；12、生产加工；13、仓储物流；99、其他
        String jzytQgStr = map.get("jzytQg") == null ? "" : map.get("jzytQg").toString();
        String[] jzytQgArr = jzytQgStr.split(",");
        String jzytQg = "";
        for (int i = 0; i < jzytQgArr.length; i++) {
            switch (jzytQgArr[i]) {
                case "1":
                    jzytQg += ",教育设施";
                    if ("是".equals(jyssQg)) {
                        jzytQg += "(中小学幼儿园教学用房及学生宿舍、食堂)";
                    }
                    break;
                case "2":
                    jzytQg += ",医疗卫生";
                    if ("是".equals(ylssQg)) {
                        jzytQg += "(具有外科手术室或急诊科的乡镇卫生院医疗用房)";
                    }
                    break;
                case "3":
                    jzytQg += ",行政办公";
                    break;
                case "4":
                    jzytQg += ",文化设施";
                    break;
                case "5":
                    jzytQg += ",养老服务";
                    break;
                case "6":
                    jzytQg += ",批发零售";
                    break;
                case "7":
                    jzytQg += ",餐饮服务";
                    break;
                case "8":
                    jzytQg += ",住宿宾馆";
                    break;
                case "9":
                    jzytQg += ",休闲娱乐";
                    break;
                case "10":
                    jzytQg += ",宗教场所";
                    break;
                case "11":
                    jzytQg += ",农贸市场";
                    break;
                case "12":
                    jzytQg += ",生产加工";
                    break;
                case "13":
                    jzytQg += ",仓储物流";
                    break;
                case "99":
                    jzytQg += "," + (map.get("qtjzytQg") == null ? "" : map.get("qtjzytQg").toString());
                    break;
                default:
            }
        }
        if (jzytQg.length() > 0) {
            jzytQg = jzytQg.substring(1);
        }
        // 是否经过安全鉴定
        String sfjgaqjdStr = map.get("sfjgaqjd") == null ? "" : map.get("sfjgaqjd").toString();
        String sfjgaqjd = "1".equals(sfjgaqjdStr) ? "是" : "0".equals(sfjgaqjdStr) ? "否" : "";
        // 是否经过安全鉴定是1时，有鉴定时间和鉴定或评定结论
        String aqjdnf = "";
        String aqjdjl = "";
        if ("1".equals(sfjgaqjdStr)) {
            // 鉴定时间
            aqjdnf = map.get("aqjdnf") == null ? "" : map.get("aqjdnf").toString();
            // 鉴定或评定结论-等级 1、A级；2、B级；3、C级；4、D级
            int aqjdjlNum = map.get("aqjdjl") == null ? -1 : Integer.parseInt(map.get("aqjdjl").toString());
            aqjdjl = "";
            switch (aqjdjlNum) {
                case 1:
                    aqjdjl = "A级";
                    break;
                case 2:
                    aqjdjl = "B级";
                    break;
                case 3:
                    aqjdjl = "C级";
                    break;
                case 4:
                    aqjdjl = "D级";
                    break;
                default:
            }
            // 鉴定或评定结论-安全性
            String jdsfaqStr = map.get("jdsfaq") == null ? "" : map.get("jdsfaq").toString();
            aqjdjl = StringUtils.isNotBlank(jdsfaqStr) ? "1".equals(jdsfaqStr) ? "安全" : "不安全" : aqjdjl;
        }

        document.open();
        WordUtil.paragrahCenter(document, title, fontHei);
        Table table = WordUtil.craeteWordTable(12, 100, 5, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1);
        table.setTop(-200);
        table.setOffset(0);
        // 未选中的复选框：□
        // 选中的复选框：☑
        WordUtil.addCell(table, "未选中的复选框：□  ", 12, 1, null, fontSongBold);
        WordUtil.addCell(table, "选中的复选框：☑  多来几个：☑ □复选框：□  ", 12, 2, null, fontSongBold);

        WordUtil.addCell(table, "第二部分：建筑信息", 12, 2, null, fontSongBold);
        WordUtil.addCell(table, "2.1 建筑层数", 2, 1, null, fontSong);
        WordUtil.addCell(table, cs + "层", 4, 1, null, fontSong);
        WordUtil.addCell(table, "2.2 建筑面积", 2, 1, null, fontSong);
        WordUtil.addCell(table, dcmj + "平方米", 4, 1, null, fontSong);
        WordUtil.addCell(table, "2.3 建造年代", 2, 1, null, fontSong);
        WordUtil.addCell(table, jzndQg, 10, 1, null, fontSong);
        WordUtil.addCell(table, "2.4 结构类型", 2, 1, null, fontSong);
        WordUtil.addCell(table, jglxQg, 10, 1, null, fontSong);
        // 第二部分序号控制
        int secondSerialNumber = 5;
        if (type == 1) {
            WordUtil.addCell(table, "2.5 建造方式", 2, 1, null, fontSong);
            WordUtil.addCell(table, jzfsQg, 10, 1, null, fontSong);
            secondSerialNumber++;
        } else if (type == 3) {
            WordUtil.addCell(table, "2.5 建造方式", 2, 1, null, fontSong);
            WordUtil.addCell(table, jzfsQg, 10, 1, null, fontSong);
            WordUtil.addCell(table, "2.6 建筑用途", 2, 1, null, fontSong);
            WordUtil.addCell(table, jzytQg, 10, 1, null, fontSong);
            secondSerialNumber += 2;
        }
        // 是否经过安全鉴定为是时，才有鉴定时间和鉴定或评定结论
        if ("是".equals(sfjgaqjd)) {
            WordUtil.addCell(table, "2." + secondSerialNumber + " 安全鉴定", 2, 2, null, fontSong);
            WordUtil.addCell(table, "2." + secondSerialNumber + ".1 是否经过安全鉴定", 4, 1, null, fontSong);
            WordUtil.addCell(table, sfjgaqjd, 2, 1, null, fontSong);
            WordUtil.addCell(table, "2." + secondSerialNumber + ".2 鉴定时间", 2, 1, null, fontSong);
            WordUtil.addCell(table, aqjdnf + "年", 2, 1, null, fontSong);
            WordUtil.addCell(table, "2." + secondSerialNumber + ".3 鉴定或评定结论", 4, 1, null, fontSong);
            WordUtil.addCell(table, aqjdjl, 6, 1, null, fontSong);
        } else {
            WordUtil.addCell(table, "2." + secondSerialNumber + " 安全鉴定", 2, 1, null, fontSong);
            WordUtil.addCell(table, "2." + secondSerialNumber + ".1 是否经过安全鉴定", 4, 1, null, fontSong);
            WordUtil.addCell(table, sfjgaqjd, 6, 1, null, fontSong);
        }

        document.add(table);
        document.close();
    }
}