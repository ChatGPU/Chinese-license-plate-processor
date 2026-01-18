from pathlib import Path
import re
import unicodedata

import pandas as pd

# =====================================================================================
# --- CONFIGURATION ZONE ---
# Modify the values in this section to adapt the script to your specific needs.
# =====================================================================================

CONFIG = {
    # The name of the column in your Excel file that contains the license plates.
    # e.g., "License Plate Number", "Plate ID", "车牌号"
    "input_column_name": "车牌号",
    # Optional aliases for auto-detecting the plate column.
    "input_column_aliases": ["车牌", "牌照", "License Plate", "Plate", "Plate Number"],
    # Optional keywords for fuzzy matching if aliases fail.
    "input_column_keywords": ["车牌", "牌照", "plate", "licenseplate"],

    # The names for the new columns that will be added to the Excel file.
    "output_province_column": "车牌归属地（省）",
    "output_city_column": "车牌归属地（市）",

    # The name of the folder where the processed files will be saved.
    "output_folder_name": "处理后表格",

    # Batch input settings.
    # "input_paths" can include files, folders, or glob patterns (e.g., "*.xlsx").
    "input_paths": ["."],
    # Whether to search subfolders of input directories.
    "recursive_search": False,
    # Whether to preserve other sheets when processing a workbook.
    "preserve_other_sheets": True,
    # Process all sheets when True, otherwise only the first sheet.
    "process_all_sheets": False,
    # Explicit sheet names to process. Leave empty to use the defaults above.
    "sheet_names": [],
    # Skip hidden files and directories (names starting with a dot).
    "skip_hidden_files": True,
    # Overwrite existing output columns if they already exist.
    "overwrite_existing_output_columns": True,
    # Skip output if no sheet contains the target column.
    "skip_files_without_column": True,
}

# =====================================================================================
# --- DATA SOURCE ---
# This dictionary maps the 2-character license plate prefix to its location.
# You can update this list if new prefixes are added or if corrections are needed.
# Format: 'Prefix': 'Province-City'
# =====================================================================================

# --- 【根据最新补充数据进行合并更新】全国车牌号开头与地区对应表 ---
PROVINCE_MAP = {
    # --- 北京市 ---
    '京A': '北京市-北京市', '京B': '北京市-北京市', '京C': '北京市-北京市', '京D': '北京市-北京市',
    '京E': '北京市-北京市', '京H': '北京市-北京市', '京J': '北京市-北京市', '京K': '北京市-北京市',
    '京L': '北京市-北京市', '京M': '北京市-北京市', '京N': '北京市-北京市', '京O': '北京市-北京市',
    '京V': '北京市-北京市', '京Y': '北京市-北京市',
    # --- 上海市 ---
    '沪A': '上海市-上海市', '沪B': '上海市-上海市', '沪C': '上海市-上海市', '沪D': '上海市-上海市',
    '沪E': '上海市-上海市', '沪F': '上海市-上海市', '沪G': '上海市-上海市', '沪H': '上海市-上海市',
    '沪J': '上海市-上海市', '沪K': '上海市-上海市', '沪L': '上海市-上海市', '沪M': '上海市-上海市',
    '沪N': '上海市-上海市', '沪R': '上海市-崇明区',
    # --- 天津市 ---
    '津A': '天津市-天津市', '津B': '天津市-天津市', '津C': '天津市-天津市', '津D': '天津市-天津市',
    '津E': '天津市-天津市', '津F': '天津市-天津市', '津G': '天津市-天津市', '津H': '天津市-天津市',
    # --- 重庆市 ---
    '渝A': '重庆市-重庆市', '渝B': '重庆市-重庆市', '渝C': '重庆市-永川区', '渝D': '重庆市-重庆市主城区',
    '渝F': '重庆市-万州区', '渝G': '重庆市-涪陵区', '渝H': '重庆市-黔江区',
    # --- 河北省 ---
    '冀A': '河北省-石家庄市', '冀B': '河北省-唐山市', '冀C': '河北省-秦皇岛市', '冀D': '河北省-邯郸市',
    '冀E': '河北省-邢台市', '冀F': '河北省-保定市', '冀G': '河北省-张家口市', '冀H': '河北省-承德市',
    '冀J': '河北省-沧州市', '冀R': '河北省-廊坊市', '冀S': '河北省-沧州市', '冀T': '河北省-衡水市',
    '冀X': '河北省-雄安新区',
    # --- 河南省 ---
    '豫A': '河南省-郑州市', '豫B': '河南省-开封市', '豫C': '河南省-洛阳市', '豫D': '河南省-平顶山市',
    '豫E': '河南省-安阳市', '豫F': '河南省-鹤壁市', '豫G': '河南省-新乡市', '豫H': '河南省-焦作市',
    '豫J': '河南省-濮阳市', '豫K': '河南省-许昌市', '豫L': '河南省-漯河市', '豫M': '河南省-三门峡市',
    '豫N': '河南省-商丘市', '豫P': '河南省-周口市', '豫Q': '河南省-驻马店市', '豫R': '河南省-南阳市',
    '豫S': '河南省-信阳市', '豫U': '河南省-济源市', '豫V': '河南省-郑州市',
    # --- 云南省 ---
    '云A': '云南省-昆明市', '云C': '云南省-昭通市', '云D': '云南省-曲靖市',
    '云E': '云南省-楚雄彝族自治州', '云F': '云南省-玉溪市', '云G': '云南省-红河哈尼族彝族自治州',
    '云H': '云南省-文山壮族苗族自治州', '云J': '云南省-思茅区', '云K': '云南省-西双版纳傣族自治州',
    '云L': '云南省-大理白族自治州', '云M': '云南省-保山市', '云N': '云南省-德宏傣族景颇族自治州',
    '云P': '云南省-丽江市', '云Q': '云南省-怒江傈僳族自治州', '云R': '云南省-迪庆藏族自治州',
    '云S': '云南省-临沧市',
    # --- 辽宁省 ---
    '辽A': '辽宁省-沈阳市', '辽B': '辽宁省-大连市', '辽C': '辽宁省-鞍山市', '辽D': '辽宁省-抚顺市',
    '辽E': '辽宁省-本溪市', '辽F': '辽宁省-丹东市', '辽G': '辽宁省-锦州市', '辽H': '辽宁省-营口市',
    '辽J': '辽宁省-阜新市', '辽K': '辽宁省-辽阳市', '辽L': '辽宁省-盘锦市', '辽M': '辽宁省-铁岭市',
    '辽N': '辽宁省-朝阳市', '辽P': '辽宁省-葫芦岛市', '辽V': '辽宁省-省直机关',
    # --- 黑龙江省 ---
    '黑A': '黑龙江省-哈尔滨市', '黑B': '黑龙江省-齐齐哈尔市', '黑C': '黑龙江省-牡丹江市',
    '黑D': '黑龙江省-佳木斯市', '黑E': '黑龙江省-大庆市', '黑F': '黑龙江省-伊春市',
    '黑G': '黑龙江省-鸡西市', '黑H': '黑龙江省-鹤岗市', '黑J': '黑龙江省-双鸭山市',
    '黑K': '黑龙江省-七台河市', '黑L': '黑龙江省-哈尔滨市(原松花江地区并入)', '黑M': '黑龙江省-绥化市',
    '黑N': '黑龙江省-黑河市', '黑P': '黑龙江省-大兴安岭地区', '黑R': '黑龙江省-农垦系统',
    # --- 湖南省 ---
    '湘A': '湖南省-长沙市', '湘B': '湖南省-株洲市', '湘C': '湖南省-湘潭市', '湘D': '湖南省-衡阳市',
    '湘E': '湖南省-邵阳市', '湘F': '湖南省-岳阳市', '湘G': '湖南省-张家界市', '湘H': '湖南省-益阳市',
    '湘J': '湖南省-常德市', '湘K': '湖南省-娄底市', '湘L': '湖南省-郴州市', '湘M': '湖南省-永州市',
    '湘N': '湖南省-怀化市', '湘U': '湖南省-湘西土家族苗族自治州',
    # --- 安徽省 ---
    '皖A': '安徽省-合肥市', '皖B': '安徽省-芜湖市', '皖C': '安徽省-蚌埠市', '皖D': '安徽省-淮南市',
    '皖E': '安徽省-马鞍山市', '皖F': '安徽省-淮北市', '皖G': '安徽省-铜陵市', '皖H': '安徽省-安庆市',
    '皖J': '安徽省-黄山市', '皖K': '安徽省-阜阳市', '皖L': '安徽省-宿州市', '皖M': '安徽省-滁州市',
    '皖N': '安徽省-六安市', '皖P': '安徽省-宣城市', '皖Q': '安徽省-巢湖市', '皖R': '安徽省-池州市',
    '皖S': '安徽省-亳州市',
    # --- 山东省 ---
    '鲁A': '山东省-济南市', '鲁B': '山东省-青岛市', '鲁C': '山东省-淄博市', '鲁D': '山东省-枣庄市',
    '鲁E': '山东省-东营市', '鲁F': '山东省-烟台市', '鲁G': '山东省-潍坊市', '鲁H': '山东省-济宁市',
    '鲁J': '山东省-泰安市', '鲁K': '山东省-威海市', '鲁L': '山东省-日照市', '鲁M': '山东省-滨州市',
    '鲁N': '山东省-德州市', '鲁P': '山东省-聊城市', '鲁Q': '山东省-临沂市', '鲁R': '山东省-菏泽市',
    '鲁S': '山东省-济南市(原莱芜市并入)', '鲁U': '山东省-青岛市', '鲁V': '山东省-潍坊市', '鲁Y': '山东省-烟台市',
    # --- 新疆维吾尔自治区 ---
    '新A': '新疆维吾尔自治区-乌鲁木齐市', '新B': '新疆维吾尔自治区-昌吉回族自治州', '新C': '新疆维吾尔自治区-石河子市',
    '新D': '新疆维吾尔自治区-奎屯市', '新E': '新疆维吾尔自治区-博尔塔拉蒙古自治州', '新F': '新疆维吾尔自治区-伊犁哈萨克自治州',
    '新G': '新疆维吾尔自治区-塔城地区', '新H': '新疆维吾尔自治区-阿勒泰地区', '新J': '新疆维吾尔自治区-克拉玛依市',
    '新K': '新疆维吾尔自治区-吐鲁番市', '新L': '新疆维吾尔自治区-哈密市', '新M': '新疆维吾尔自治区-巴音郭愣蒙古自治州',
    '新N': '新疆维吾尔自治区-阿克苏地区', '新P': '新疆维吾尔自治区-克孜勒苏柯尔克孜自治州', '新Q': '新疆维吾尔自治区-喀什地区',
    '新R': '新疆维吾尔自治区-和田地区',
    # --- 江苏省 ---
    '苏A': '江苏省-南京市', '苏B': '江苏省-无锡市', '苏C': '江苏省-徐州市', '苏D': '江苏省-常州市',
    '苏E': '江苏省-苏州市', '苏F': '江苏省-南通市', '苏G': '江苏省-连云港市', '苏H': '江苏省-淮安市',
    '苏J': '江苏省-盐城市', '苏K': '江苏省-扬州市', '苏L': '江苏省-镇江市', '苏M': '江苏省-泰州市',
    '苏N': '江苏省-宿迁市', '苏U': '江苏省-苏州市',
    # --- 浙江省 ---
    '浙A': '浙江省-杭州市', '浙B': '浙江省-宁波市', '浙C': '浙江省-温州市', '浙D': '浙江省-绍兴市',
    '浙E': '浙江省-湖州市', '浙F': '浙江省-嘉兴市', '浙G': '浙江省-金华市', '浙H': '浙江省-衢州市',
    '浙J': '浙江省-台州市', '浙K': '浙江省-丽水市', '浙L': '浙江省-舟山市',
    # --- 江西省 ---
    '赣A': '江西省-南昌市', '赣B': '江西省-赣州市', '赣C': '江西省-宜春市', '赣D': '江西省-吉安市',
    '赣E': '江西省-上饶市', '赣F': '江西省-抚州市', '赣G': '江西省-九江市', '赣H': '江西省-景德镇市',
    '赣J': '江西省-萍乡市', '赣K': '江西省-新余市', '赣L': '江西省-鹰潭市', '赣M': '江西省-南昌市',
    # --- 湖北省 ---
    '鄂A': '湖北省-武汉市', '鄂B': '湖北省-黄石市', '鄂C': '湖北省-十堰市', '鄂D': '湖北省-荆州市',
    '鄂E': '湖北省-宜昌市', '鄂F': '湖北省-襄阳市', '鄂G': '湖北省-鄂州市', '鄂H': '湖北省-荆门市',
    '鄂J': '湖北省-黄冈市', '鄂K': '湖北省-孝感市', '鄂L': '湖北省-咸宁市', '鄂M': '湖北省-仙桃市',
    '鄂N': '湖北省-潜江市', '鄂P': '湖北省-神农架林区', '鄂Q': '湖北省-恩施土家族苗族自治州',
    '鄂R': '湖北省-天门市', '鄂S': '湖北省-随州市', '鄂W': '湖北省-武汉市',
    # --- 广西壮族自治区 ---
    '桂A': '广西壮族自治区-南宁市', '桂B': '广西壮族自治区-柳州市', '桂C': '广西壮族自治区-桂林市',
    '桂D': '广西壮族自治区-梧州市', '桂E': '广西壮族自治区-北海市', '桂F': '广西壮族自治区-崇左市',
    '桂G': '广西壮族自治区-来宾市', '桂H': '广西壮族自治区-桂林市', '桂J': '广西壮族自治区-贺州市',
    '桂K': '广西壮族自治区-玉林市', '桂L': '广西壮族自治区-百色市', '桂M': '广西壮族自治区-河池市',
    '桂N': '广西壮族自治区-钦州市', '桂P': '广西壮族自治区-防城港市', '桂R': '广西壮族自治区-贵港市',
    # --- 甘肃省 ---
    '甘A': '甘肃省-兰州市', '甘B': '甘肃省-嘉峪关市', '甘C': '甘肃省-金昌市', '甘D': '甘肃省-白银市',
    '甘E': '甘肃省-天水市', '甘F': '甘肃省-酒泉市', '甘G': '甘肃省-张掖市', '甘H': '甘肃省-武威市',
    '甘J': '甘肃省-定西市', '甘K': '甘肃省-陇南市', '甘L': '甘肃省-平凉市', '甘M': '甘肃省-庆阳市',
    '甘N': '甘肃省-临夏回族自治州', '甘P': '甘肃省-甘南藏族自治州',
    # --- 山西省 ---
    '晋A': '山西省-太原市', '晋B': '山西省-大同市', '晋C': '山西省-阳泉市', '晋D': '山西省-长治市',
    '晋E': '山西省-晋城市', '晋F': '山西省-朔州市', '晋H': '山西省-忻州市', '晋J': '山西省-吕梁市',
    '晋K': '山西省-晋中市', '晋L': '山西省-临汾市', '晋M': '山西省-运城市',
    # --- 内蒙古自治区 ---
    '蒙A': '内蒙古自治区-呼和浩特市', '蒙B': '内蒙古自治区-包头市', '蒙C': '内蒙古自治区-乌海市',
    '蒙D': '内蒙古自治区-赤峰市', '蒙E': '内蒙古自治区-呼伦贝尔市', '蒙F': '内蒙古自治区-兴安盟',
    '蒙G': '内蒙古自治区-通辽市', '蒙H': '内蒙古自治区-锡林郭勒盟', '蒙J': '内蒙古自治区-乌兰察布市',
    '蒙K': '内蒙古自治区-鄂尔多斯市', '蒙L': '内蒙古自治区-巴彦淖尔市', '蒙M': '内蒙古自治区-阿拉善盟',
    # --- 陕西省 ---
    '陕A': '陕西省-西安市', '陕B': '陕西省-铜川市', '陕C': '陕西省-宝鸡市', '陕D': '陕西省-咸阳市',
    '陕E': '陕西省-渭南市', '陕F': '陕西省-汉中市', '陕G': '陕西省-安康市', '陕H': '陕西省-商洛市',
    '陕J': '陕西省-延安市', '陕K': '陕西省-榆林市', '陕U': '陕西省-西安市', '陕V': '陕西省-杨凌区',
    # --- 吉林省 ---
    '吉A': '吉林省-长春市', '吉B': '吉林省-吉林市', '吉C': '吉林省-四平市', '吉D': '吉林省-辽源市',
    '吉E': '吉林省-通化市', '吉F': '吉林省-白山市', '吉G': '吉林省-白城市', '吉H': '吉林省-延边朝鲜族自治州',
    '吉J': '吉林省-松原市', '吉K': '吉林省-长白朝鲜族自治县',
    # --- 福建省 ---
    '闽A': '福建省-福州市', '闽B': '福建省-莆田市', '闽C': '福建省-泉州市', '闽D': '福建省-厦门市',
    '闽E': '福建省-漳州市', '闽F': '福建省-龙岩市', '闽G': '福建省-三明市', '闽H': '福建省-南平市',
    '闽J': '福建省-宁德市', '闽K': '福建省-省直系统',
    # --- 贵州省 ---
    '贵A': '贵州省-贵阳市', '贵B': '贵州省-六盘水市', '贵C': '贵州省-遵义市', '贵D': '贵州省-铜仁市',
    '贵E': '贵州省-黔西南布依族苗族自治州', '贵F': '贵州省-毕节市', '贵G': '贵州省-安顺市',
    '贵H': '贵州省-黔东南苗族侗族自治州', '贵J': '贵州省-黔南布依族苗族自治州',
    # --- 广东省 ---
    '粤A': '广东省-广州市', '粤B': '广东省-深圳市', '粤C': '广东省-珠海市', '粤D': '广东省-汕头市',
    '粤E': '广东省-佛山市', '粤F': '广东省-韶关市', '粤G': '广东省-湛江市', '粤H': '广东省-肇庆市',
    '粤J': '广东省-江门市', '粤K': '广东省-茂名市', '粤L': '广东省-惠州市', '粤M': '广东省-梅州市',
    '粤N': '广东省-汕尾市', '粤P': '广东省-河源市', '粤Q': '广东省-阳江市', '粤R': '广东省-清远市',
    '粤S': '广东省-东莞市', '粤T': '广东省-中山市', '粤U': '广东省-潮州市', '粤V': '广东省-揭阳市',
    '粤W': '广东省-云浮市', '粤X': '广东省-顺德区', '粤Y': '广东省-南海区', '粤Z': '广东省-港澳进入内地车辆',
    # --- 四川省 ---
    '川A': '四川省-成都市', '川B': '四川省-绵阳市', '川C': '四川省-自贡市', '川D': '四川省-攀枝花市',
    '川E': '四川省-泸州市', '川F': '四川省-德阳市', '川G': '四川省-成都市', '川H': '四川省-广元市',
    '川J': '四川省-遂宁市', '川K': '四川省-内江市', '川L': '四川省-乐山市', '川M': '四川省-资阳市',
    '川Q': '四川省-宜宾市', '川R': '四川省-南充市', '川S': '四川省-达州市', '川T': '四川省-雅安市',
    '川U': '四川省-阿坝藏族羌族自治州', '川V': '四川省-甘孜藏族自治州', '川W': '四川省-凉山彝族自治州',
    '川X': '四川省-广安市', '川Y': '四川省-巴中市', '川Z': '四川省-眉山市',
    # --- 青海省 ---
    '青A': '青海省-西宁市', '青B': '青海省-海东市', '青C': '青海省-海北藏族自治州', '青D': '青海省-黄南藏族自治州',
    '青E': '青海省-海南藏族自治州', '青F': '青海省-果洛藏族自治州', '青G': '青海省-玉树藏族自治州',
    '青H': '青海省-海西蒙古族藏族自治州',
    # --- 西藏自治区 ---
    '藏A': '西藏自治区-拉萨市', '藏B': '西藏自治区-昌都市', '藏C': '西藏自治区-山南市', '藏D': '西藏自治区-日喀则市',
    '藏E': '西藏自治区-那曲市', '藏F': '西藏自治区-阿里地区', '藏G': '西藏自治区-林芝市',
    '藏H': '西藏自治区-驻四川省天全县车辆管理所', '藏J': '西藏自治区-驻青海省格尔木市车辆管理所',
    # --- 海南省 ---
    '琼A': '海南省-海口市', '琼B': '海南省-三亚市', '琼C': '海南省-琼海市', '琼D': '海南省-五指山市',
    '琼E': '海南省-洋浦开发区', '琼F': '海南省-儋州市',
    # --- 宁夏回族自治区 ---
    '宁A': '宁夏回族自治区-银川市', '宁B': '宁夏回族自治区-石嘴山市', '宁C': '宁夏回族自治区-吴忠市',
    '宁D': '宁夏回族自治区-固原市', '宁E': '宁夏回族自治区-中卫市',
}


# =====================================================================================
# --- CORE LOGIC ---
# No need to modify anything below this line for normal use.
# =====================================================================================

SUPPORTED_EXTENSIONS = {".xlsx", ".xls"}
PLATE_CLEAN_PATTERN = re.compile(r"[^0-9A-Za-z\u4e00-\u9fff]")


def normalize_column_name(name):
    return re.sub(r"[\s\-_]+", "", str(name)).lower()


def resolve_input_column(columns, config):
    normalized_map = {}
    for col in columns:
        normalized_map.setdefault(normalize_column_name(col), []).append(col)

    candidates = [config.get("input_column_name")] + list(config.get("input_column_aliases", []))
    for candidate in candidates:
        if not candidate:
            continue
        normalized = normalize_column_name(candidate)
        if normalized in normalized_map:
            matches = normalized_map[normalized]
            if len(matches) > 1:
                print(f"  -> WARNING: Multiple columns match '{candidate}', using '{matches[0]}'.")
            elif candidate != config.get("input_column_name"):
                print(f"  -> INFO: Column '{matches[0]}' matched via alias '{candidate}'.")
            return matches[0]

    keywords = [normalize_column_name(keyword) for keyword in config.get("input_column_keywords", []) if keyword]
    if keywords:
        keyword_matches = [
            col
            for col in columns
            if any(keyword in normalize_column_name(col) for keyword in keywords)
        ]
        if len(keyword_matches) == 1:
            print(f"  -> INFO: Column '{keyword_matches[0]}' matched via keyword.")
            return keyword_matches[0]
        if len(keyword_matches) > 1:
            print("  -> WARNING: Multiple columns matched keywords; please set 'input_column_name'.")

    return None


def normalize_plate_series(series):
    cleaned = series.astype("string")
    cleaned = cleaned.map(
        lambda value: unicodedata.normalize("NFKC", value) if isinstance(value, str) else value
    )
    cleaned = cleaned.str.strip()
    cleaned = cleaned.str.replace(PLATE_CLEAN_PATTERN, "", regex=True)
    cleaned = cleaned.str.upper()
    return cleaned


def build_location_columns(series):
    cleaned = normalize_plate_series(series)
    lengths = cleaned.str.len().fillna(0)
    prefix = cleaned.str.slice(0, 2)
    prefix = prefix.where(lengths >= 2, "")
    location = prefix.map(PROVINCE_MAP).fillna("未知-未知")
    parts = location.str.split("-", n=1, expand=True).reindex(columns=[0, 1])
    province = parts[0].fillna("未知")
    city = parts[1].fillna(parts[0]).fillna("未知")
    return province, city


def make_unique_column_name(desired_name, existing_columns, reserved_columns, allow_overwrite):
    if allow_overwrite and desired_name not in reserved_columns:
        return desired_name

    candidate = desired_name
    if candidate in existing_columns or candidate in reserved_columns:
        suffix = " (new)"
        candidate = f"{desired_name}{suffix}"
        counter = 2
        while candidate in existing_columns or candidate in reserved_columns:
            candidate = f"{desired_name}{suffix} {counter}"
            counter += 1
    return candidate


def reorder_columns(df, input_col, new_cols):
    cols = list(df.columns)
    for col in new_cols:
        cols = [existing for existing in cols if existing != col]
    try:
        insert_at = cols.index(input_col)
    except ValueError:
        insert_at = len(cols)
    final_cols = cols[:insert_at] + list(new_cols) + cols[insert_at:]
    return df[final_cols]


def get_reader_engine(file_path):
    suffix = file_path.suffix.lower()
    if suffix == ".xlsx":
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    return None


def resolve_output_path(file_path, output_dir, current_dir):
    suffix = file_path.suffix.lower()
    output_suffix = suffix
    writer_engine = "openpyxl"
    warning = None

    if suffix == ".xls":
        try:
            import xlwt  # noqa: F401
        except Exception:
            output_suffix = ".xlsx"
            writer_engine = "openpyxl"
            warning = "xlwt is not installed; saving .xls as .xlsx instead."
        else:
            writer_engine = "xlwt"
    elif suffix == ".xlsx":
        writer_engine = "openpyxl"
    else:
        output_suffix = ".xlsx"
        warning = f"Unsupported extension '{suffix}', saving as .xlsx."

    try:
        relative_path = file_path.resolve().relative_to(current_dir.resolve())
        output_subdir = relative_path.parent
    except ValueError:
        safe_parts = [
            part for part in file_path.resolve().parent.parts if part not in (file_path.anchor, "")
        ]
        output_subdir = Path("external") / Path(*safe_parts)

    output_folder = output_dir / output_subdir
    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / f"{file_path.stem}{output_suffix}"
    return output_path, writer_engine, warning


def is_supported_file(file_path, output_dir, skip_hidden):
    if skip_hidden and any(part.startswith(".") for part in file_path.parts):
        return False

    try:
        resolved_path = file_path.resolve()
    except FileNotFoundError:
        resolved_path = file_path

    if output_dir and (output_dir == resolved_path or output_dir in resolved_path.parents):
        return False

    return file_path.suffix.lower() in SUPPORTED_EXTENSIONS


def collect_excel_files(input_paths, output_dir, recursive, skip_hidden):
    current_dir = Path.cwd()
    output_dir_resolved = output_dir.resolve()
    seen = set()
    files = []

    for raw_path in input_paths:
        if raw_path is None:
            continue
        raw_path = str(raw_path)
        if not raw_path:
            continue
        if any(char in raw_path for char in ["*", "?", "["]):
            path = Path(raw_path)
            if path.is_absolute():
                base = Path(path.anchor)
                pattern = str(path.relative_to(path.anchor))
                candidates = list(base.glob(pattern))
            else:
                candidates = list(current_dir.glob(raw_path))
        else:
            path = Path(raw_path)
            if not path.is_absolute():
                path = current_dir / path
            candidates = [path]

        for candidate in candidates:
            if not candidate.exists():
                print(f"  -> WARNING: Input path not found: {candidate}")
                continue
            if candidate.is_dir():
                iterator = candidate.rglob("*") if recursive else candidate.glob("*")
                for file_path in iterator:
                    if not file_path.is_file():
                        continue
                    if not is_supported_file(file_path, output_dir_resolved, skip_hidden):
                        continue
                    resolved = file_path.resolve()
                    if resolved in seen:
                        continue
                    seen.add(resolved)
                    files.append(file_path)
            elif candidate.is_file():
                if not is_supported_file(candidate, output_dir_resolved, skip_hidden):
                    continue
                resolved = candidate.resolve()
                if resolved in seen:
                    continue
                seen.add(resolved)
                files.append(candidate)

    return sorted(files, key=lambda path: str(path).lower())


def select_sheet_names(all_sheet_names, config):
    explicit = [name for name in config.get("sheet_names") or [] if name]
    if explicit:
        matched = [name for name in explicit if name in all_sheet_names]
        missing = [name for name in explicit if name not in all_sheet_names]
        if missing:
            print(f"  -> WARNING: Sheet(s) not found: {', '.join(missing)}")
        return matched

    if config.get("process_all_sheets"):
        return list(all_sheet_names)

    return all_sheet_names[:1]


def format_display_path(path, base_dir):
    try:
        return str(path.relative_to(base_dir))
    except ValueError:
        return str(path)


def process_dataframe(df, config, context_prefix=""):
    input_col = resolve_input_column(df.columns, config)
    if not input_col:
        message = f"{context_prefix} SKIPPED: Column '{config['input_column_name']}' not found."
        print(message.strip())
        return df, False

    existing_columns = list(df.columns)
    reserved_columns = {input_col}
    allow_overwrite = bool(config.get("overwrite_existing_output_columns", True))

    prov_col = make_unique_column_name(
        config["output_province_column"],
        existing_columns,
        reserved_columns,
        allow_overwrite,
    )
    if prov_col != config["output_province_column"]:
        print(
            f"{context_prefix} WARNING: Output column '{config['output_province_column']}' exists; "
            f"using '{prov_col}'."
        )

    reserved_columns.add(prov_col)
    city_col = make_unique_column_name(
        config["output_city_column"],
        existing_columns + [prov_col],
        reserved_columns,
        allow_overwrite,
    )
    if city_col != config["output_city_column"]:
        print(
            f"{context_prefix} WARNING: Output column '{config['output_city_column']}' exists; "
            f"using '{city_col}'."
        )

    province, city = build_location_columns(df[input_col])
    df[prov_col] = province
    df[city_col] = city
    df = reorder_columns(df, input_col, [prov_col, city_col])
    return df, True


def write_excel_file(output_path, sheets, writer_engine):
    with pd.ExcelWriter(output_path, engine=writer_engine) as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def process_excel_file(file_path, output_dir, current_dir, config):
    engine = get_reader_engine(file_path)
    if engine is None:
        print(f"  -> SKIPPED: Unsupported file type '{file_path.suffix}'.")
        return "skipped"

    try:
        with pd.ExcelFile(file_path, engine=engine) as excel_file:
            sheet_names = excel_file.sheet_names
            if not sheet_names:
                print("  -> SKIPPED: No sheets found.")
                return "skipped"

            target_sheet_names = select_sheet_names(sheet_names, config)
            if not target_sheet_names:
                print("  -> SKIPPED: No matching sheets found.")
                return "skipped"

            output_sheets = {}
            processed_any = False
            for sheet_name in sheet_names:
                if not config.get("preserve_other_sheets", True) and sheet_name not in target_sheet_names:
                    continue
                df = excel_file.parse(sheet_name)
                if sheet_name in target_sheet_names:
                    context = f"  [Sheet: {sheet_name}]"
                    df, processed = process_dataframe(df, config, context)
                    processed_any = processed_any or processed
                output_sheets[sheet_name] = df

            if config.get("skip_files_without_column", True) and not processed_any:
                print("  -> SKIPPED: Target column not found in selected sheets.")
                return "skipped"

            output_path, writer_engine, warning = resolve_output_path(
                file_path,
                output_dir,
                current_dir,
            )
            if warning:
                print(f"  -> WARNING: {warning}")

            write_excel_file(output_path, output_sheets, writer_engine)
            display_path = format_display_path(output_path, current_dir)
            print(f"  -> SUCCESS: Saved to '{display_path}'")
            return "processed"
    except ImportError as e:
        print(f"  -> ERROR: Missing dependency for '{file_path.name}': {e}")
        return "error"
    except Exception as e:
        print(f"  -> ERROR: An unexpected error occurred: {e}")
        return "error"


def process_license_plates_in_directory():
    """
    Main function to find, process, and save Excel files.
    """
    current_dir = Path.cwd()
    output_dir = current_dir / CONFIG["output_folder_name"]
    output_dir.mkdir(parents=True, exist_ok=True)

    input_paths = CONFIG.get("input_paths", ["."])
    if isinstance(input_paths, (str, Path)):
        input_paths = [str(input_paths)]

    print("Starting script...")
    print(f"Current directory: {current_dir}")
    print(f"Input paths: {input_paths}")
    print(f"Recursive search: {CONFIG.get('recursive_search', False)}")
    print(f"Processed files will be saved to: {output_dir}")
    print("-" * 30)

    excel_files = collect_excel_files(
        input_paths,
        output_dir,
        recursive=CONFIG.get("recursive_search", False),
        skip_hidden=CONFIG.get("skip_hidden_files", True),
    )

    if not excel_files:
        print("No Excel files (.xlsx or .xls) found in the specified paths.")
        return

    processed_count = 0
    skipped_count = 0
    error_count = 0

    for file_path in excel_files:
        display_path = format_display_path(file_path, current_dir)
        print(f"Processing: {display_path}")
        result = process_excel_file(file_path, output_dir, current_dir, CONFIG)
        if result == "processed":
            processed_count += 1
        elif result == "skipped":
            skipped_count += 1
        else:
            error_count += 1

    print("-" * 30)
    print(
        "All files processed. "
        f"{processed_count} updated, {skipped_count} skipped, {error_count} error(s)."
    )

if __name__ == '__main__':
    # --- How to use ---
    # 1. Install required libraries:
    #    pip install pandas openpyxl xlrd
    #    (Optional for .xls output) pip install xlwt
    #
    # 2. Place this script in the same folder as your Excel files.
    #
    # 3. (Optional) Adjust the settings in the CONFIGURATION ZONE at the top.
    #
    # 4. Run the script from your terminal:
    #    python your_script_name.py
    process_license_plates_in_directory()
