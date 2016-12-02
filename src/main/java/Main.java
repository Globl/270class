package main.java;

import com.sun.org.apache.regexp.internal.RE;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

    public static void main(String[] args) {

        System.out.println("Hello World!");

        String classNameStr = "不做“豆芽菜”——塑形健身,天文观测,欧洲文化游,汽车美容,汽车驾驶与保养,招贴画设计,汽车使用常识,汽车文化,时装设计,创意与制作之虚拟机器人,海报制作,摄影作品欣赏与拍摄技巧,摄影艺术,工艺蛋壳画,“书”之美，“法”自然,创意与数字美术,装饰艺术,创意手工,模拟联合国,中式烹饪技艺,微视频之微中有戏,旅游十八计 ,西点制作,高中生涯规划辅导,3D动画基础,用数学的眼光看世界,心语.心晴,导游英语口语,谁说你不懂电路,民谣弹唱初级教程,跨文化交际教程,创意与写作,国学研究三部曲,中国古代散文专题阅读欣赏,中外文学名著及其影片欣赏,服装模特形体训练,唐诗中的地理,消费化学,导游词的表达技巧与模拟讲解,手绘DIY,装饰画制作,素描基础绘画——石膏几何体画法,雕塑制作与欣赏,图形创意,剪纸艺术传承与创新,走向生活,现代社交礼仪,现代陶艺教程,做一个有能力的人——职业素养训练,校园常见植物的鉴赏与种植,互动媒体技术,微电子商务宝典,Flash 动漫影片设计,智能机器人制作与竞赛,游戏中的数学,几何画板与迷你数学实验（上）,篮球规则与裁判法,国际跳棋,生活中的物理,交通故事  非常物理 ,高中生自信心培育,走进神秘的心理世界,高中生心理课堂,电子小制作,旅游英语,微电影鉴赏与制作,趣味戏曲,校园话剧表演,音乐与戏剧表演,原声影视英语配音与赏析,法律与生活,高中女生性健康必备,探秘海洋,生活与地理,驴友行走浙江,大话旅游,生存与环境,舌尖上的化学,钳工小制作,竹刻艺术,木工基础教程,标志创意设计,中国古代文史常识,动画鉴赏与设计,水彩画艺术,现代家庭装修之电路安装,艺术插花,茶艺基础,家用电器,航空航天知识及模型制作,西点烘焙,能力生物学,职业生涯规划,PS多彩的世界,故事趣说数学逻辑,数学的魅力,素描基础,排球入门,物理建模,汽车中的物理,趣味色彩心理学,化学与烘焙,高中英语报刊赏析,高中英语美文欣赏读本,鲁迅小说欣赏,论语与哲学,生活法律导航,中国酒文化概要,茶艺鉴赏,数码照片的PHOTOSHOP后期制作,创业导论,快乐生物,中国旅游文化地图—民俗篇,观景入门,天文气象,军事武器中的化学,当舌尖与化学碰撞,环境 生活 化学,化学智慧与化学史,食用化学,学做环境检测员,化学与社会,工业产品设计,电子线路焊接,翻簧竹雕艺术,趣味电子技术,简明中国文化与中国哲学,走进二十四史,中国园林的瑰宝-圆明园,创意漫画,集邮——从小方寸中见大世界,越窑青瓷,创意软陶,沙画艺术,刮版画,鼠标下的艺术,我的理财小课堂,海洋渔业常识,室内装饰品设计入门,家庭园艺,走进美容,花卉栽培与花卉鉴赏,常见花卉与栽培,DIY十字绣,发酵与美食,中学生风光摄影入门,品读茶文化,课余小花匠,东海海洋物产,三维建模与室内设计,《毕业季》微电影后期设计与制作,校园电子商务,透过数学看经济,生活中的数学视角,MATLAB入门,健美操选修课程,网球学与练,透过电影看物理,玩转魔方,戏剧表演,健康从心开始,情绪妙管家,追寻幸福人生,FLASH矢量图小动画设计,微电子创意制作,航模制作及飞行技术,生活与地理,食品中的化学,中草药标本制作,中国饮食文化的地域差异,形体舞蹈,高中音乐鉴赏与创作综合课程,英美文化之旅,教你读懂鲁迅,小小说课堂,中国古代史史料学初步,我们一起来作文,中国民间文学趣谈,汉字魅力,红楼寻梦,古诗文中的生命超越哲学,“字”里春秋,生活中的写作,分享我的散文,西方经典建筑的文化解读,《红楼梦》女性形象赏析,新闻评论阅读与写作,摹形传神      千载如生 ——《史记》人物刻画艺术鉴赏 ,美国短篇小说名作赏析,《三国演义》专题讲座,赏漫画.学政治,身边的经济现象,现代中国的国际关系,饮食文化与健康饮食,现代礼仪风尚,创意与天文观测及摄影,海鲜饮食与营养健康,中医研究,世界国家地理,居住环境与地理,海岛世界,《现代自然地理学》专题讲座,世界自然遗产选析,化学趣味探索实验,饮水思源――从喝中学化学,见证生活处处皆化学,化学与文学,家庭作坊中的化学制造,实用皮革皮毛加工技术(上),机械设计初步,汽车维修与保养,走进物流,电子产品制作,中国古代文物欣赏与仿制,走近电工技术,古代中外文化交流专题史,中国传统文化习俗十讲,钓鱼岛的那些事儿,生活化的历史-走近浙江非物质文化遗产,书法中的历史,美国总统评说,手绘POP和海报制作,动态沙画表演,当水笔遇上草书,铅笔风景速写,素描,数码照片影楼级技能速成,绿色果蔬空中生态种植,中国结艺,走近西饼师行业,投资理财入门,中国风味面点,头饰小制作,高中生物实验拓展,校园植物学探索,高中生涯规划——体验式学习,安卓平台下的创意程序设计,“微”我风采——微电影拍摄与编辑,数列的诱惑,几何画板与数学实验,数学探究案例赏析,趣味物理课堂,会声会影打造校园视频故事,健康身心，幸福一生-高中女生健康教育读本,生物显微摄影,校园气象观测与物候,我爱编歌曲,英语诗歌赏析入门,赏经典动画   学地道口语 ,走进世界著名大学 (A Visit to the Famous Universities around the World),新闻评论写作,诗文中的盛世——读着唐诗学唐史,当代热点环境问题,微信文艺范,生活要懂点经济学,健康与幸福,风云浙商,浙江美丽的非物质文化遗产,化学与服饰,信息安全与密码,走近物联网技术,微视界摄影,个人电脑的安全与急救,人体健康与高中生物学,photoshop照片处理与创意设计,小故事大数学,国学经典选讲,从《新闻联播》看外交知识,我眼中的中国建筑,演讲•主持•辩论";
        String[] classNameArr = classNameStr.split(",");
        System.out.println(classNameArr.length);
        Map<String,String > map = new HashMap<String, String>();
        for (String className : classNameArr){
            if (map.containsKey(className)){
                System.out.println("重复的课程名:"+className);
                continue;
            }
            map.put(className,className);
        }
        System.out.println("去重后的map大小:"+map.size());
        try {
            InputStream is = new FileInputStream("270classes.xlsx");
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            List<ClassItem> classItemList = new ArrayList<ClassItem>();
            Map<String,List<ClassItem>> classItemMap = new HashMap<String, List<ClassItem>>();
            for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
                if (xssfSheet == null) {
                    continue;
                }
                // Read the Row
                for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    if (xssfRow != null) {
                        ClassItem classItem = new ClassItem();
                        XSSFCell classNameXss = xssfRow.getCell(1);
                        XSSFCell catalogNameXss = xssfRow.getCell(2);
                        XSSFCell fileNameXss = xssfRow.getCell(3);
                        XSSFCell filePathXss = xssfRow.getCell(4);
                        classItem.setFilePath(filePathXss.getStringCellValue());
                        classItem.setClassName(classNameXss.getStringCellValue());
                        classItem.setCatalogName(catalogNameXss.getStringCellValue());
                        classItem.setFileName(getValue(fileNameXss));
                        classItemList.add(classItem);
                        if (!classItemMap.containsKey(classItem.getClassName())){
                            classItemMap.put(classItem.getClassName(),new ArrayList<ClassItem>());
                        }else{
                            System.out.println(classItem.getClassName()+" 重复了");
                        }
                        classItemMap.get(classItem.getClassName()).add(classItem);
                    }
                }
            }
            System.out.println(classItemMap.size());

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getValue(XSSFCell fileNameXss) {
        if (fileNameXss.getCellType() == XSSFCell.CELL_TYPE_NUMERIC){
            return fileNameXss.getNumericCellValue()+"";
        }else {
            return fileNameXss.getStringCellValue();
        }
    }


}
