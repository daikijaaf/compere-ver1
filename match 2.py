import pandas as pd
file_path = "" #file path 入力
df_leadgen= pd.read_excel(file_path, sheet_name=0, index_col=0)
df_uber = pd.read_excel(file_path, sheet_name=1)
# print(df_leadgen.head)
print(df_uber['name'].head())
print(df_leadgen.keys())
print(df_uber.keys())
# 参考サイト https://akizora.tech/python-address-divide-4231
# 住所分割  prefecture_jp, city_jp, af_city(市区町村のあとの住所)
import re
df_leadgen['prefecture_jp'] = ""
df_leadgen['city_jp'] = ""
df_leadgen['af_city'] = ""
for index, row in df_leadgen.iterrows():
    f_address = row['Formatted Restaurant Address']
    address_list = f_address.split("\n")
    address = ""
    if len(address_list) == 3:
        address = address_list[2]
    else:
        address = address_list[0]
    matches = re.match(r'(...??[都道府県])((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下松|岩国|田川|大村)市|.+?郡(?:玉村|大町|.+?)[町村]|.+?市.+?区|.+?[市区町村])(.+)' , address)
    if matches is not None:
        row['prefecture_jp'] = matches[1]
        row['city_jp'] = matches[2]
        row['af_city'] = matches[3]
    else:
        # 都道府県がない場合
        matches = re.match(r'((?:旭川|伊達|石狩|盛岡|奥州|田村|南相馬|那須塩原|東村山|武蔵村山|羽村|十日町|上越|富山|野々市|大町|蒲郡|四日市|姫路|大和郡山|廿日市|下松|岩国|田川|大村)市|.+?郡(?:玉村|大町|.+?)[町村]|.+?市.+?区|.+?[市区町村])(.+)' , address)
        if matches is not None:
            row['city_jp'] = matches[1]
            row['af_city'] = matches[2]
        else:
            pass
    # 元のデータフレームを更新する
    df_leadgen.loc[index] = row

pip install  Levenshtein
# 文字列の類似度計算 参考サイト http://pixelbeat.jp/text-matching-3-approach-with-python/#toc_id_2
import Levenshtein  # 文字列の類似度計算

# 元のデータフレームは残しておいて、編集用にコピーする
df_leadgen_af = df_leadgen.copy()
df_leadgen_af['match_group'] = ""
df_leadgen_af['uber_name'] = ""
df_leadgen_af['rating'] = ""
df_leadgen_af['number_of_reviews'] = ""
near_match_list = []
total_row = df_leadgen_af['Account Name'].count()
counter = 0
for index, leadgen_row in df_leadgen_af.iterrows():
    counter = counter + 1
    account_name = str(leadgen_row['Account Name'])
    match_flag = False
    print("処理中{}/{}: {}, {}".format(str(counter),str(total_row),index,leadgen_row['Account Name']))
    for index, uber_row in df_uber.iterrows():
        uber_name = str(uber_row['name'])
        jaro_dist = Levenshtein.jaro_winkler(account_name, uber_name)
        # 店舗名が完全一致している場合
        if jaro_dist == 1.0:
            print("   店舗名一致",account_name,uber_name,jaro_dist)
            match_flag = True
            # 店舗名と住所と電話番号を比較する
            if leadgen_row['prefecture_jp'] == uber_row['prefecture_jp'] and leadgen_row['city_jp'] == uber_row['city_jp']:
                if leadgen_row['Phone'] == uber_row['phone_number']:
                    leadgen_row['match_group'] = "match_name_city_phone"
                else:
                    leadgen_row['match_group'] = "match_name_city"
            # 店舗名と電話番号で比較する(住所なし)
            else:
                if leadgen_row['Phone'] == uber_row['phone_number']:
                    leadgen_row['match_group'] = "match_name_phone"
                else:
                    leadgen_row['match_group'] = "match_name"
        # 0.9 以上だと店舗名一致していて、支店名が異なるケースなどがマッチする(0.93~0.95あたりのスコアになる)
        elif jaro_dist > 0.9:
            match_flag = True
            print("   類似度0.9以上一致:",account_name,uber_name,jaro_dist)
            near_match_list.append([account_name,uber_name,jaro_dist])
            # 店舗名類似と住所と電話番号を比較する
            if leadgen_row['prefecture_jp'] == uber_row['prefecture_jp'] and leadgen_row['city_jp'] == uber_row['city_jp']:
                if leadgen_row['Phone'] == uber_row['phone_number']:
                    leadgen_row['match_group'] = "match_nearname_city_phone"
                else:
                    leadgen_row['match_group'] = "match_nearname_city"
            # 店舗名類似と電話番号で比較する(住所なし)
            else:
                if leadgen_row['Phone'] == uber_row['phone_number']:
                    leadgen_row['match_group'] = "match_nearname_phone"
                else:
                    leadgen_row['match_group'] = "match_nearname"
                    leadgen_row['uber_name'] = "__NEARMATCH__:" + uber_row['name'] 
                    match_flag = False
        if match_flag is True:
            # 元のデータフレームを更新する
            leadgen_row['uber_name'] = str(uber_row['name'])
            leadgen_row['rating'] = str(uber_row['rating'])
            leadgen_row['number_of_reviews'] = str(uber_row['number_of_reviews'])
            df_leadgen_af.loc[index] = leadgen_row
            # 一致しているデータがある場合、# 処理中の比較元の行(leadgenシートの行)と比較先(uberシート)の比較を終了し、
            # 次のleadgenシートの行の比較処理をおこなう
            break


# マッチ結果の確認
df_leadgen_af[df_leadgen_af['match_group'] != ""]


# エクセルに出力
df_leadgen_af.to_excel('xxx', sheet_name='xxx')# add file & sheet name



