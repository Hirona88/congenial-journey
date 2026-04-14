import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ===== Sheet 1: メインリスト =====
ws1 = wb.active
ws1.title = "海外バイヤーリスト"

# カラー定義
HEADER_BG   = "1F4E79"   # 濃紺
HEADER_FG   = "FFFFFF"
S3_BG       = "FF0000"   # ★★★ 赤
S2_BG       = "FF9900"   # ★★  オレンジ
S1_BG       = "92D050"   # ★   緑
ALT_BG      = "EBF3FB"   # 薄青（交互行）
WHITE       = "FFFFFF"

def hdr_fill(color): return PatternFill("solid", fgColor=color)
def cell_fill(color): return PatternFill("solid", fgColor=color)

thin = Side(style="thin", color="CCCCCC")
med  = Side(style="medium", color="888888")
thin_border = Border(left=thin, right=thin, top=thin, bottom=thin)
med_border  = Border(left=med,  right=med,  top=med,  bottom=med)

# ヘッダー行
headers = [
    "No", "優先度", "会社名", "国", "地域",
    "ウェブサイト", "メールアドレス", "電話番号", "WhatsApp",
    "Facebook", "SNS(その他)", "BR380JG取扱実績", "備考"
]
ws1.append(headers)
for col, _ in enumerate(headers, 1):
    c = ws1.cell(row=1, column=col)
    c.fill   = hdr_fill(HEADER_BG)
    c.font   = Font(color=HEADER_FG, bold=True, size=10)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = thin_border
ws1.row_dimensions[1].height = 30

# データ
data = [
    # No, 優先度, 会社名, 国, 地域, Web, Email, Tel, WA, FB, SNS他, 実績, 備考
    (1,  "★★★", "Stevens Group NZ",                          "ニュージーランド", "オセアニア",
     "https://www.stevensgroup.co.nz",  "sales@stevensgroup.co.nz",
     "+64 9 275 0443\nLionel Stanners: +64 21 278 0738\nDan Bartley: +64 21 278 7456",
     "", "https://www.facebook.com/p/Stevens-Group-NZ-100090097132297/",
     "LinkedIn あり", "あり（掲載実績）",
     "NZ最大級の破砕・ふるい機械ディーラー。BR380JGを過去に販売リスト掲載。中古機も扱う。"),

    (2,  "★★★", "Anhui Zhaomi Engineering Machinery",         "中国（安徽省合肥市）", "東アジア",
     "https://mpmachinery.en.alibaba.com", "Alibaba メッセージ",
     "", "+86 193 0560 5637",
     "https://www.facebook.com/p/Anhui-zhaomi-engineering-machinery-co-ltd-100094293633667/",
     "", "あり（販売実績）",
     "日本からKomatsu/CAT/日立中古重機を輸出専門。20年以上の実績。BR380JGをMachinio・MachineryTraderに複数掲載。最有力ターゲット。"),

    (3,  "★★★", "Winwin Used Machinery",                      "韓国", "東アジア",
     "http://www.winwinused.com", "simson@winwinused.com",
     "+82-2-553-7007", "",
     "", "Alibaba掲載 / Skype: eboyshop1", "あり（Alibaba販売実績）",
     "コマツBR380JGをAlibabaで過去に販売。韓国発の中古重機輸出業者。ショベル・クラッシャー・ローラー等取扱。"),

    (4,  "★★★", "Miller Plant & Equipment",                   "オーストラリア（ビクトリア州）", "オセアニア",
     "http://millermachinery.com.au", "Webフォームより",
     "+61 3 9314 0744", "",
     "https://www.facebook.com/millerauspty/", "", "あり（Machines4u掲載）",
     "日本・米国から中古建機を輸入して販売。コマツBR380取扱実績あり（Machines4u掲載）。"),

    (5,  "★★",  "Al Marwan Machinery",                        "UAE", "中東",
     "https://almarwan.com", "Webフォームより",
     "", "", "", "公式SNSあり", "なし",
     "UAE・サウジ・オマーンでKomatsu/CAT/日立/Volvo等の中古機を販売。Powerscreen正規代理店。ジョー・コーン・インパクトクラッシャー取扱。"),

    (6,  "★★",  "Galadari Trucks & Heavy Equipment",          "UAE", "中東",
     "https://www.galadarigthe.com", "Webフォームより",
     "", "", "", "", "なし",
     "コマツと45年のパートナーシップ。UAE最大規模の建機ディーラーの一つ。Metso製クラッシャーも扱う。"),

    (7,  "★★",  "Saleh Al-Juhani Sons (Al-Juhni)",            "サウジアラビア", "中東",
     "https://www.al-juhni.com/en/crushers", "Webフォームより",
     "", "", "", "", "なし",
     "サウジの骨材・採石業界トップ企業。メディナほかに複数の破砕機拠点。高い購買力。"),

    (8,  "★★",  "Jaffer Brothers Private Limited",             "パキスタン", "南アジア",
     "https://www.jaffer.com/machinery", "Webフォームより",
     "+92 21 111-527-527 (カラチ)\n+92 51 111-527-527 (イスラマバード)\n+92 42 111-527-527 (ラホール)",
     "", "https://www.facebook.com/jaffergroup/", "", "なし",
     "パキスタンにコマツを初めて導入した老舗代理店（1948年創業）。採石・建設機械全般。コマツ中東正規ディーラー。"),

    (9,  "★★",  "International Crusher Solutions (EA) Ltd",   "ケニア（ナイロビ）", "アフリカ",
     "https://www.internationalcrushersolutions.co.ke", "enquiries@instcrush.com",
     "+254 718 774 324", "",
     "https://www.facebook.com/InternationalCrusherSolutions/", "LinkedIn あり", "なし",
     "東アフリカ最大の採石機械・クラッシャー専門業者。70台以上の在庫。タンザニア公認代理店。新旧クラッシャー取扱。"),

    (10, "★★",  "Pilot Crushtec International (Pty) Ltd",     "南アフリカ", "アフリカ",
     "https://www.pilotcrushtec.com", "Webフォームより",
     "+27 11 842 5600", "", "", "公式SNSあり", "なし",
     "南アフリカ最大の移動式・半固定式破砕機サプライヤー。Metso公認代理店。モバイルジョークラッシャー取扱。"),

    (11, "★★",  "Z&M Machinery",                              "中国（上海）", "東アジア",
     "Machineryline掲載", "Machinerylineより連絡",
     "", "", "", "", "あり（掲載実績）",
     "MachinerylineにKomatsu BR380クラッシャーを掲載中。2年以上の出品実績。"),

    (12, "★★",  "2CMACHINERY Ltd",                            "中国", "東アジア",
     "MachineryTrader Asia掲載", "MachineryTrader Asiaより連絡",
     "", "", "", "", "あり",
     "MachineryTrader Asiaにて中古Komatsuクラッシャーを出品。"),

    (13, "★★",  "Maxima Machineries Inc.",                    "フィリピン", "東南アジア",
     "https://maxima.com.ph", "Webフォームより",
     "", "", "https://www.facebook.com/maximamachineriesinc/", "", "なし",
     "丸紅グループ。フィリピン唯一のコマツ正規代理店。コマツ建設・鉱山機械全般を扱う。"),

    (14, "★★",  "Kraftvélar",                                 "アイスランド", "欧州",
     "", "", "", "",
     "Facebook（Komatsu Europe投稿で確認）", "", "新品導入実績",
     "Facebook上でKomatsu BR380JG-3の新品導入実績あり（Komatsu Europe公式投稿より確認）。中古需要が見込まれる。"),

    (15, "★★",  "Crusher Equipment Africa (CE Africa)",        "南アフリカ", "アフリカ",
     "https://ce-africa.com", "Webフォームより",
     "", "", "", "", "なし",
     "RUBBLE MASTER公認代理店。ジョー・コーン・ジャイラトリークラッシャー等のメンテ・販売。"),

    (16, "★★",  "Komatsu Australia",                          "オーストラリア", "オセアニア",
     "https://www.komatsu.com.au/equipment/crushers", "Webフォームより",
     "", "", "", "公式SNSあり", "あり",
     "コマツ公式ディーラー。新旧BR380JG-1E0取扱。中古機の需要旺盛。"),

    (17, "★",   "Komatsu Philippines Corporation (KPC)",       "フィリピン", "東南アジア",
     "https://www.kpc.komatsu", "Webフォームより",
     "", "", "", "", "なし",
     "コマツ子会社。販売窓口はMaximaへ。"),

    (18, "★",   "DIMO (David Pieris Motor Co.)",               "スリランカ", "南アジア",
     "https://www.dimolanka.com", "Webフォームより",
     "", "", "", "", "なし",
     "コマツと50年以上のパートナーシップ。建設・採石用機械を扱う。"),

    (19, "★",   "Komatsu Sales (Cambodia) / CMED Group",       "カンボジア", "東南アジア",
     "https://cmedgp.com", "Webフォームより",
     "+855 23 900 589", "", "", "", "なし",
     "カンボジアのコマツ正規代理店。"),

    (20, "★",   "Shanghai AKE Used Machinery",                 "中国（上海）", "東アジア",
     "http://www.secondexcavator.com", "Webフォームより",
     "", "", "", "", "なし",
     "日本からの中古建機輸出。コマツ・日立等取扱。"),

    (21, "★",   "Equipment Africa Ltd",                        "ガーナ", "アフリカ",
     "https://www.equipmentafrica.com", "Webフォームより",
     "", "", "", "", "なし",
     "西アフリカ向け国際的建機ソリューションプロバイダー。ガーナ発で複数国でレンタル・売却。"),

    (22, "★",   "Komatsu Poland",                              "ポーランド", "欧州",
     "https://komatsupoland.pl/en/", "Webフォームより",
     "", "", "", "", "なし",
     "ポーランド唯一のコマツ正規代理店。新旧コマツ機械の販売・整備・レンタル。"),
]

priority_colors = {"★★★": S3_BG, "★★": S2_BG, "★": S1_BG}

for i, row in enumerate(data, 2):
    ws1.append(list(row))
    bg = ALT_BG if i % 2 == 0 else WHITE
    prio = row[1]
    for col in range(1, len(headers) + 1):
        c = ws1.cell(row=i, column=col)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = thin_border
        if col == 2:  # 優先度列
            c.fill = cell_fill(priority_colors.get(prio, WHITE))
            c.font = Font(bold=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.fill = cell_fill(bg)
    ws1.row_dimensions[i].height = 50

# 列幅設定
col_widths = [5, 8, 28, 18, 12, 38, 30, 32, 22, 45, 18, 16, 50]
for col, w in enumerate(col_widths, 1):
    ws1.column_dimensions[get_column_letter(col)].width = w

# 先頭行を固定
ws1.freeze_panes = "A2"
ws1.auto_filter.ref = ws1.dimensions

# ===== Sheet 2: プラットフォームリスト =====
ws2 = wb.create_sheet("プラットフォーム一覧")

plat_headers = ["No", "プラットフォーム名", "URL", "特記事項"]
ws2.append(plat_headers)
for col in range(1, 5):
    c = ws2.cell(row=1, column=col)
    c.fill = hdr_fill(HEADER_BG)
    c.font = Font(color=HEADER_FG, bold=True, size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border
ws2.row_dimensions[1].height = 25

platforms = [
    (1,  "Machinio",              "https://www.machinio.com/komatsu/br380/crushers-screening-plants",         "BR380JG-3を複数掲載中。世界最大級の中古建機マーケット。"),
    (2,  "Machineryline（US/UK/JP版）","https://machineryline.com",                                           "BR380JGを€330,000〜で掲載。欧州・米国バイヤー接点。"),
    (3,  "MachineryZone",         "https://www.machineryzone.com/used/crusher/1/3090/komatsu/br-380-jg.html", "欧州中心。BR380JG多数掲載。"),
    (4,  "MachineryTrader / MachineryTrader Asia","https://www.machinerytrader.com",                          "北米・アジア向け。BR380JGリスト有り。"),
    (5,  "Equippo",               "https://www.equippo.com/en/product/crusher/komatsu-br380jg-1e0/",          "BR380JG-1E0を€97,500でフランスから掲載実績。"),
    (6,  "BIGLEMON",              "https://biglemon.kenkey.jp/en",                                             "日本最大級の中古建機ネットオークション。海外バイヤー向け英語対応あり。"),
    (7,  "HeavyMart",             "https://www.heavymart.com",                                                 "アジア最大級の中古建機マーケット。インドネシア・マレーシア・タイ等でコマツ掲載多数。"),
    (8,  "Toku World（日本）",    "https://toku-world.com/en/stock/recycling_machine/mobile-jaw-crusher/maker/komatsu/","日本発の中古建機輸出。BR380JG-3在庫あり（BKJ041）。世界向け輸出実績。"),
    (9,  "Auto Link Holdings（日本）","https://www.autolink.co.jp/type/machinery/mobile+crusher/",             "日本から中古モバイルクラッシャーを輸出。"),
    (10, "ESTEC Trade（日本）",   "https://estec-trade.com",                                                   "フィリピン・インドネシアほか世界向けに日本中古建機を輸出。BR380JG-1の実績あり。"),
    (11, "Alibaba（Komatsu BR380JG関連）","https://www.alibaba.com/showroom/komatsu-japanese-crusher.html",   "Winwin Used Machinery（韓国）がBR380JGを過去出品。中国発業者も多数出品中。"),
    (12, "Ritchie Bros. / IronPlanet","https://www.rbauction.com",                                             "国際オークション。コマツ機多数。欧米・中東バイヤーが集まる。"),
    (13, "Mascus",                "https://www.mascus.com",                                                    "欧州・南アフリカでコマツ中古機多数掲載。"),
    (14, "Growth Power（日本）",  "https://growthpower.jp/English/other-construction/komatsu/BR380JG-3",       "日本からのBR380JG-3輸出実績あり。"),
]

for i, row in enumerate(platforms, 2):
    ws2.append(list(row))
    bg = ALT_BG if i % 2 == 0 else WHITE
    for col in range(1, 5):
        c = ws2.cell(row=i, column=col)
        c.fill = cell_fill(bg)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = thin_border
    ws2.row_dimensions[i].height = 35

ws2.column_dimensions["A"].width = 5
ws2.column_dimensions["B"].width = 28
ws2.column_dimensions["C"].width = 65
ws2.column_dimensions["D"].width = 55
ws2.freeze_panes = "A2"

# ===== Sheet 3: SNSアカウント一覧 =====
ws3 = wb.create_sheet("SNSアカウント一覧")

sns_headers = ["No", "プラットフォーム", "アカウント名 / 投稿内容", "URL / 備考"]
ws3.append(sns_headers)
for col in range(1, 5):
    c = ws3.cell(row=1, column=col)
    c.fill = hdr_fill(HEADER_BG)
    c.font = Font(color=HEADER_FG, bold=True, size=10)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border
ws3.row_dimensions[1].height = 25

sns_data = [
    (1,  "Facebook", "Anhui Zhaomi Engineering Machinery（中国・合肥）",          "https://www.facebook.com/p/Anhui-zhaomi-engineering-machinery-co-ltd-100094293633667/"),
    (2,  "Facebook", "Stevens Group NZ（NZ・クラッシャー専門）",                   "https://www.facebook.com/p/Stevens-Group-NZ-100090097132297/"),
    (3,  "Facebook", "Maxima Machineries Inc.（フィリピン・コマツ代理店）",        "https://www.facebook.com/maximamachineriesinc/"),
    (4,  "Facebook", "Komatsu Europe（BR380JG-3紹介動画・bauma2022公式）",         "https://www.facebook.com/KomatsuEurope/"),
    (5,  "Facebook", "International Crusher Solutions（ケニア・東アフリカ）",       "https://www.facebook.com/InternationalCrusherSolutions/"),
    (6,  "Facebook", "Jaffer Group（パキスタン・コマツ代理店）",                    "https://www.facebook.com/jaffergroup/"),
    (7,  "Facebook", "満天建機（日本・BR380JGメンテ完了投稿あり）",                 "mantenkenki ページより（日本国内業者・輸出業者紹介元として活用可）"),
    (8,  "Facebook グループ", "Mobile Crusher BR380JG-1 KOMATSU（投稿あり）",      "グループID: 1485495605090156"),
    (9,  "Facebook グループ", "Jaw crusher Komatsu BR380JG for sale 投稿（2015年式）","グループID: 495760541639862"),
    (10, "Instagram", "2018年式 BR380JG-3（32時間）出品投稿あり",                   "https://www.instagram.com/p/C3Z-x0LBBm8/"),
    (11, "Instagram", "KomatsuNZ（ニュージーランド・公式）",                        "@komatsunewzealand"),
    (12, "LinkedIn",  "International Crusher Solutions Ltd",                        "https://uk.linkedin.com/company/instant-crusher-spares-ltd"),
    (13, "LinkedIn",  "Stevens Group NZ",                                           "https://nz.linkedin.com/company/stevens-group-nz"),
]

for i, row in enumerate(sns_data, 2):
    ws3.append(list(row))
    bg = ALT_BG if i % 2 == 0 else WHITE
    for col in range(1, 5):
        c = ws3.cell(row=i, column=col)
        c.fill = cell_fill(bg)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.border = thin_border
    ws3.row_dimensions[i].height = 35

ws3.column_dimensions["A"].width = 5
ws3.column_dimensions["B"].width = 18
ws3.column_dimensions["C"].width = 50
ws3.column_dimensions["D"].width = 65
ws3.freeze_panes = "A2"

# ===== 保存 =====
out = "/home/user/congenial-journey/overseas_buyers_BR380JG3.xlsx"
wb.save(out)
print(f"Saved: {out}")
