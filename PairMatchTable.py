import sys
import openpyxl
import datetime
import random
from collections import namedtuple
import zenhan
import math

# 性別でペアを組むならTrue
seibetsu_flag = False

# ペアを組むならTrue
pair_kotei_flag = True

# 乱数での棋力変動幅
random_size = 3

# タイトルの行位置
title_row = 1

# 参加者リストの開始行位置
start_sankasha_row = 3

# 参加者リストの名前列位置
sankasha_name_col = 1

# 対戦回の開始列位置
taisen_start_col = 4

# 勝ち数の列位置
WIN_col = 0

# SOSの列位置
SOS_col = 0

# SOSOSの列位置
SOSOS_col = 0

# 順位の列位置
JYUNI_col = 0

# 参加者のデータ構造体定義
taisensha_info = namedtuple("taisensha_info", "row pair_no random_row name kiryoku seibetsu score taisen_rireki sos sosos jyuni special")
# 参加者の対戦履歴のデータ構造体定義
taisen_rireki = namedtuple("taisen_rireki", "no name pair_name taisensha_name1 taisensha_name2 kekka")
# 対戦相手のデータ構造体定義
taisen_aite_info = namedtuple("taisen_aite_info", "taisensha_info taikyoku_su index")

# 棋力を取得(5D=5,1D=1,1K=0,2K=-1)
def get_kiryoku(dat):
    kiryoku = 0
    if dat[1:2] == 'D':
        kiryoku = int(dat.replace('D', ''))
    else:
        kiryoku = 1 - int(dat.replace('K', ''))
    return kiryoku

# 対戦未決定リスト取得
def get_taisen_mikettei_list(taisensha_info_rec, pair_info_rec, taisensha_info_list, taisenNo, step):

    mikettei_list = []

    for rec in taisensha_info_list:
        if rec.name == taisensha_info_rec.name:
            continue
        if rec.name == pair_info_rec.name:
            continue
        # 既に対戦番号(1～)以上の対局を実施した人は対象外
        if len(rec.taisen_rireki) >= taisenNo:
            continue
        # 過去に対戦した人は対象外
        if step == 1:
            past_flag = False
            for rireki in taisensha_info_rec.taisen_rireki:
                if rireki.taisensha_name1 == rec.name:
                    past_flag = True
                    break
                if rireki.taisensha_name2 == rec.name:
                    past_flag = True
                    break
            if past_flag:
                continue
        # 相手は同性でなければ対象外
        if seibetsu_flag:
            if taisensha_info_rec.seibetsu != rec.seibetsu:
                continue
        # それ以外を候補とする
        mikettei_list.append(rec)

    return mikettei_list

# ペア未決定リスト取得
def get_pair_mikettei_list(taisensha_info_rec, taisensha_info_list, taisenNo, step):

    mikettei_list = []

    for rec in taisensha_info_list:
        if rec.name == taisensha_info_rec.name:
            continue
        # 既に対戦番号(1～)以上の対局を実施した人は対象外
        if len(rec.taisen_rireki) >= taisenNo:
            continue
        # 過去にペアを組んだ人は対象外
        if step == 1:
            past_pairs_flag = False
            for rireki in taisensha_info_rec.taisen_rireki:
                if rireki.pair_name == rec.name:
                    past_pairs_flag = True
                    break
            if past_pairs_flag:
                continue
        # 同性は対象外
        if seibetsu_flag:
            if taisensha_info_rec.seibetsu == rec.seibetsu:
                continue
        # それ以外を候補とする
        mikettei_list.append(rec)

    return mikettei_list

# 対戦相手のペア未決定リスト取得
def get_pair_mikettei_list2(taisensha_info_rec, pair_info_rec, taisen_aite_info_rec, taisensha_info_list, taisenNo, step):

    mikettei_list = []

    for rec in taisensha_info_list:
        if rec.name == taisensha_info_rec.name:
            continue
        if rec.name == pair_info_rec.name:
            continue
        if rec.name == taisen_aite_info_rec.name:
            continue
        # 既に対戦番号(1～)以上の対局を実施した人は対象外
        if len(rec.taisen_rireki) >= taisenNo:
            continue
        # 過去にペアを組んだ人は対象外
        if step == 1:
            past_pairs_flag = False
            for rireki in taisen_aite_info_rec.taisen_rireki:
                if rireki.pair_name == rec.name:
                    past_pairs_flag = True
                    break
            if past_pairs_flag:
                continue
        # 同性は対象外
        if seibetsu_flag:
            if taisensha_info_rec.seibetsu == rec.seibetsu:
                continue
        # それ以外を候補とする
        mikettei_list.append(rec)

    return mikettei_list

# 対戦者を決定
def get_taisen_kettei(taisensha_info_rec, pair_info_rec, taisensha_info_list, taisenNo, step):

    # 対戦者未決定リスト取得
    mikettei_list = get_taisen_mikettei_list(taisensha_info_rec, pair_info_rec, taisensha_info_list, taisenNo, step)

    # 対戦者なし
    if len(mikettei_list) == 0:
        return None

    mikettei_list2 = []

    for index, aite_info in enumerate(mikettei_list):
        # 過去対戦数を取得
        taikyoku_su = get_taikyoku_su(taisensha_info_rec.taisen_rireki, aite_info.name)
        mikettei_list2.append(taisen_aite_info(aite_info, taikyoku_su, index))

    # 対戦回数・登録順にする
    mikettei_list2 = sorted(mikettei_list2, key=lambda x: (x.taikyoku_su, x.index))

    return mikettei_list2[0].taisensha_info

# 固定ペアを取得
def get_kotei_pair(taisensha_info_rec, taisensha_info_list):
    for rec in taisensha_info_list:
        if rec.name == taisensha_info_rec.name:
            continue
        if rec.pair_no == taisensha_info_rec.pair_no:
            return rec
    return None

# ペアを決定
def get_pair_kettei(taisensha_info_rec, taisensha_info_list, taisenNo, step):

    if pair_kotei_flag:
        return get_kotei_pair(taisensha_info_rec, taisensha_info_list)        

    # ペア未決定リスト取得
    mikettei_list = get_pair_mikettei_list(taisensha_info_rec, taisensha_info_list, taisenNo, step)

    # 対戦者なし
    if len(mikettei_list) == 0:
        return None

    return mikettei_list[len(mikettei_list) - 1]

# 対戦相手のペアを決定
def get_pair_kettei2(taisensha_info_rec, pair_info_rec, taisen_aite_info_rec, taisensha_info_list, taisenNo, step):

    if pair_kotei_flag:
        return get_kotei_pair(taisen_aite_info_rec, taisensha_info_list)        

    # ペア未決定リスト取得
    mikettei_list = get_pair_mikettei_list2(taisensha_info_rec, pair_info_rec, taisen_aite_info_rec, taisensha_info_list, taisenNo, step)

    # 対戦者なし
    if len(mikettei_list) == 0:
        return None

    return mikettei_list[len(mikettei_list) - 1]

# スコアを取得
def get_score(taisen_rireki_list):

    score = 0
    for rec in taisen_rireki_list:
        if (rec.kekka == '〇'):
            score += 1
    return score

# 過去の対局数
def get_taikyoku_su(taisen_rireki_list, taisen_aite_name):
    taikyoku_su = 0
    for rec in taisen_rireki_list:
        if rec.taisensha_name1 == taisen_aite_name:
            taikyoku_su += 1
        if rec.taisensha_name2 == taisen_aite_name:
            taikyoku_su += 1
    return taikyoku_su

# 対戦相手の対戦者情報を取得
def get_aite_info(taisensha_info_list, name):
    for rec in taisensha_info_list:
        if rec.name == name:
            return rec
    return None

# SOSを計算
def get_sos(taisensha_info_list, taisensha_info_rec):
    sos = 0
    for rec in taisensha_info_rec.taisen_rireki:
        aite_info = get_aite_info(taisensha_info_list, rec.taisensha_name1)
        if aite_info != None:
            score = get_score(aite_info.taisen_rireki)
            sos += score
        aite_info = get_aite_info(taisensha_info_list, rec.taisensha_name2)
        if aite_info != None:
            score = get_score(aite_info.taisen_rireki)
            sos += score
    return sos

# SOSOSを計算
def get_sosos(taisensha_info_list, taisensha_info_rec):
    sosos = 0
    for rec in taisensha_info_rec.taisen_rireki:
        aite_info = get_aite_info(taisensha_info_list, rec.taisensha_name1)
        if aite_info != None:
            sos = get_sos(taisensha_info_list, aite_info)
            sosos += sos
        aite_info = get_aite_info(taisensha_info_list, rec.taisensha_name2)
        if aite_info != None:
            sos = get_sos(taisensha_info_list, aite_info)
            sosos += sos
    return sosos

# 成績の列位置を取得
def result_col_info(sheet):
    global WIN_col
    global SOS_col
    global SOSOS_col
    global JYUNI_col
    for col in range(1, sheet.max_column + 1):
        title = sheet.cell(title_row, col).value
        if title == '勝ち数':
            WIN_col = col
        if title == 'SOS':
            SOS_col = col
        if title == 'SOSOS':
            SOSOS_col = col
        if title == '順位':
            JYUNI_col = col

# エクセルから参加者と過去の対戦情報を読み取り
def read_excel(sheet, taisenNo, taisensha_info_list):

    for row in range(start_sankasha_row, sheet.max_row + 1):

        # 対局者名を取得
        name = sheet.cell(row, sankasha_name_col).value
        if name != None:

            pair_no = math.floor((row - start_sankasha_row) / 2 + 1)

            random_row = random.randint(-random_size, random_size)

            # 性別を取得
            seibetsu = sheet.cell(row, sankasha_name_col + 1).value

            # 棋力を取得
            kiryoku = get_kiryoku(sheet.cell(row, sankasha_name_col + 2).value)

            special = 0

            # 対戦履歴情報取得
            taisen_rireki_info_list = []

            for taisenNo_loop in range(1, taisenNo + 1):
                # ペアの名前、対戦者の名前と結果を取得
                col = taisen_start_col + (taisenNo_loop - 1) * 4
                pair_name = sheet.cell(row, col).value
                kekka = sheet.cell(row, col + 3).value
                if kekka != None:
                    # 対戦者履歴に追加
                    if pair_name == None:
                        # 不戦勝
                        special = 1
                        taisen_rireki_info = taisen_rireki(no = taisenNo_loop, name = name, pair_name = None, taisensha_name1 = None, taisensha_name2 = None, kekka = kekka)
                        taisen_rireki_info_list.append(taisen_rireki_info)
                    else:
                        taisensha_names = sheet.cell(row, col + 1).value.split('、')
                        taisensha_name1 = taisensha_names[0]
                        taisensha_name2 = taisensha_names[1]
                        taisen_rireki_info = taisen_rireki(no = taisenNo_loop, name = name, pair_name = pair_name, taisensha_name1 = taisensha_name1, taisensha_name2 = taisensha_name2, kekka = kekka)
                        taisen_rireki_info_list.append(taisen_rireki_info)

            # 参加者の勝ち星を取得
            score = get_score(taisen_rireki_info_list)

            # 参加者リストに追加
            taisensha_info_list.append(taisensha_info(row = row, pair_no = pair_no, random_row = random_row, name = name, kiryoku = kiryoku, seibetsu = seibetsu, score = score, taisen_rireki = taisen_rireki_info_list, sos = 0, sosos = 0, jyuni = 0, special = special))

# 過去の対戦情報の矛盾をチェック
def check_taisen_rireki(taisensha_info_list, last_taisen_no):
    error_flag = False
    for rec in taisensha_info_list:
        for i, senreki in enumerate(rec.taisen_rireki):
            if i < last_taisen_no:
                aite_info = get_aite_info(taisensha_info_list, senreki.taisensha_name1)
                if aite_info == None:
                    # 不戦勝などは対局者情報が取得できない。
                    continue
                if len(aite_info.taisen_rireki) > i:
                    aite_senreki = aite_info.taisen_rireki[i]
                    if senreki.kekka == aite_senreki.kekka:
                        print('{}回戦の {} vs {} 戦の結果が両者同じです。'.format(i+1, rec.name, aite_info.name))
                        error_flag = True
    return not error_flag

# ハンディキャップ取得
def get_handycap(taisensha_info_list, name1, name2, name3, name4):

    kiryoku1 = (get_aite_info(taisensha_info_list, name1).kiryoku + 
                get_aite_info(taisensha_info_list, name2).kiryoku) / 2

    kiryoku2 = (get_aite_info(taisensha_info_list, name3).kiryoku + 
                get_aite_info(taisensha_info_list, name4).kiryoku) / 2

    handy_diff = kiryoku1 - kiryoku2

    msg = ''
    if handy_diff == 0:
        msg = 'コミ６目'
    else:
        point = (handy_diff / 0.5)
        if point > 0:
            if point >= 19:
                point = 18
            point -= 1
            oki_ishi = point // 2 + 1
            komi = 0
            if (point % 2) == 1:
                komi = -6
            msg = '向 '
            if oki_ishi > 1:
                msg = msg + zenhan.h2z(str.format('{:.1g}', oki_ishi)) + '子'
            if komi == 0:
                msg = msg + 'コミなし'
            else:
                msg = msg + '逆コミ６目'
        else:
            point = point * -1
            if point >= 19:
                point = 18
            point -= 1
            oki_ishi = point // 2 + 1
            komi = 0
            if (point % 2) == 1:
                komi = -6
            msg = ''
            if oki_ishi > 1:
                msg = msg + zenhan.h2z(str.format('{:.1g}', oki_ishi)) + '子'
            if komi == 0:
                msg = msg + 'コミなし'
            else:
                msg = msg + '逆コミ６目'

    return '[' + str(handy_diff) + '] ' + msg

# 対局者決定
def player_decision(taisenNo, execel_file_name, save_execel_file_name):

    wb = openpyxl.load_workbook(execel_file_name)
    sheet = wb.active

    taisensha_info_list = []

    # エクセルから参加者と過去の対戦情報を読み取り
    read_excel(sheet, taisenNo, taisensha_info_list)

    # 過去の対戦情報の矛盾をチェック
    if check_taisen_rireki(taisensha_info_list, taisenNo - 1) == False:
        return

    # スコア・棋力（±３）・登録順にする
    taisensha_info_list = sorted(taisensha_info_list, key=lambda x: (x.special, x.score, x.kiryoku + x.random_row, x.row * -1), reverse=True)

    # 対戦の組み合わせを計算
    for i, rec in enumerate(taisensha_info_list):

        # 既に組み合わせされていたらスキップ
        if len(rec.taisen_rireki) >= taisenNo:
            continue

        # ペアを決定
        pair_rec = get_pair_kettei(rec, taisensha_info_list, taisenNo, 1)
        if pair_rec == None:
            pair_rec = get_pair_kettei(rec, taisensha_info_list, taisenNo, 2)

        if pair_rec == None:

            # 不戦勝扱いにする
            # 対戦リストに登録
            taisen_rireki_info = taisen_rireki(no = taisenNo, name = rec.name, pair_name = None, taisensha_name1 = '不戦勝', taisensha_name2 = None, kekka = '〇')
            rec.taisen_rireki.append(taisen_rireki_info)

        else:

            # 対戦者を決定
            aite_info = get_taisen_kettei(rec, pair_rec, taisensha_info_list, taisenNo, 1)
            if aite_info == None:
                aite_info = get_taisen_kettei(rec, pair_rec, taisensha_info_list, taisenNo, 2)

            # 対戦者なし
            if aite_info == None:

                # 不戦勝扱いにする
                # 対戦リストに登録
                taisen_rireki_info = taisen_rireki(no = taisenNo, name = rec.name, pair_name = None, taisensha_name1 = '不戦勝', taisensha_name2 = None, kekka = '〇')
                rec.taisen_rireki.append(taisen_rireki_info)

            else:

                # 対戦相手のペアを決定
                aite_pair_rec = get_pair_kettei2(rec, pair_rec, aite_info, taisensha_info_list, taisenNo, 1)
                if aite_pair_rec == None:
                    aite_pair_rec = get_pair_kettei2(rec, pair_rec, aite_info, taisensha_info_list, taisenNo, 2)

                if aite_pair_rec == None:

                    # 不戦勝扱いにする
                    # 対戦リストに登録
                    taisen_rireki_info = taisen_rireki(no = taisenNo, name = rec.name, pair_name = None, taisensha_name1 = '不戦勝', taisensha_name2 = None, kekka = '〇')
                    rec.taisen_rireki.append(taisen_rireki_info)

                else:

                    # 対戦リストに登録
                    taisen_rireki_info = taisen_rireki(no = taisenNo, name = rec.name, pair_name = pair_rec.name, taisensha_name1 = aite_info.name, taisensha_name2 = aite_pair_rec.name, kekka = None)
                    rec.taisen_rireki.append(taisen_rireki_info)

                    # ペアの対局リストに登録
                    taisen_rireki_info = taisen_rireki(no = taisenNo, name = pair_rec.name, pair_name = rec.name, taisensha_name1 = aite_info.name, taisensha_name2 = aite_pair_rec.name, kekka = None)
                    pair_rec.taisen_rireki.append(taisen_rireki_info)

                    # 対戦相手１の対局リストに登録
                    taisen_rireki_info = taisen_rireki(no = taisenNo, name = aite_info.name, pair_name = aite_pair_rec.name, taisensha_name1 = rec.name, taisensha_name2 = pair_rec.name, kekka = None)
                    aite_info.taisen_rireki.append(taisen_rireki_info)

                    # 対戦相手２の対局リストに登録
                    taisen_rireki_info = taisen_rireki(no = taisenNo, name = aite_pair_rec.name, pair_name = aite_info.name, taisensha_name1 = rec.name, taisensha_name2 = pair_rec.name, kekka = None)
                    aite_pair_rec.taisen_rireki.append(taisen_rireki_info)

    # エクセルにペアと対戦者名とハンディを書き込み
    for rec in taisensha_info_list:
        if len(rec.taisen_rireki) >= taisenNo:
            senreki = rec.taisen_rireki[taisenNo - 1]
            col = taisen_start_col + (taisenNo - 1) * 4
            # ペア名
            sheet.cell(rec.row, col).value = senreki.pair_name
            # 対戦相手
            if senreki.taisensha_name2 != None:
                sheet.cell(rec.row, col + 1).value = senreki.taisensha_name1 + '、' + senreki.taisensha_name2
            else:
                sheet.cell(rec.row, col + 1).value = senreki.taisensha_name1
            # ハンディ
            if senreki.taisensha_name2 != None:
                sheet.cell(rec.row, col + 2).value = get_handycap(taisensha_info_list, senreki.name, senreki.pair_name, senreki.taisensha_name1, senreki.taisensha_name2)
            # 成績(不戦勝の場合)
            if senreki.kekka != None:
                sheet.cell(rec.row, col + 3).value = senreki.kekka

    wb.save(save_execel_file_name)

# 成績決定
def write_result(execel_file_name, save_execel_file_name):

    wb = openpyxl.load_workbook(execel_file_name)
    sheet = wb.active

    # 成績の列位置を取得
    result_col_info(sheet)

    taisenNo = int((WIN_col - taisen_start_col) / 4)

    taisensha_info_list = []

    # エクセルから参加者と過去の対戦情報を読み取り
    read_excel(sheet, taisenNo, taisensha_info_list)

    # 過去の対戦情報の矛盾をチェック
    if check_taisen_rireki(taisensha_info_list, taisenNo) == False:
        return

    # SOSを計算
    taisensha_info_list_wk = []
    for rec in taisensha_info_list:
        # SOS
        sos = get_sos(taisensha_info_list, rec)
        # SOSOS
        sosos = get_sosos(taisensha_info_list, rec)
        new_rec = rec._replace(sos = sos, sosos = sosos)
        taisensha_info_list_wk.append(new_rec)
    taisensha_info_list = taisensha_info_list_wk

    # スコア・SOS・SOSOS順にする
    taisensha_info_list = sorted(taisensha_info_list, key=lambda x: (x.score, x.sos, x.sosos), reverse=True)

    # 順位を計算
    taisensha_info_list_wk = []
    current_jyuni = 1
    jyuni_count = 1
    old_rec = None
    for rec in taisensha_info_list:
        if old_rec != None:
            if old_rec.score != rec.score or old_rec.sos != rec.sos or old_rec.sosos != rec.sosos:
                if pair_kotei_flag:
                    current_jyuni += 1
                else:
                    current_jyuni += jyuni_count
                jyuni_count = 1
            else:
                jyuni_count += 1
        jyuni = current_jyuni
        old_rec = rec
        new_rec = rec._replace(jyuni = jyuni)
        taisensha_info_list_wk.append(new_rec)
    taisensha_info_list = taisensha_info_list_wk

    # 結果を書き込み
    for rec in taisensha_info_list:
        # 勝ち数
        sheet.cell(rec.row, WIN_col).value = rec.score
        # SOS
        sheet.cell(rec.row, SOS_col).value = rec.sos
        # SOSOS
        sheet.cell(rec.row, SOSOS_col).value = rec.sosos
        # 順位
        sheet.cell(rec.row, JYUNI_col).value = rec.jyuni

    wb.save(save_execel_file_name)

if __name__ == "__main__":

    excel_file_name = 'ペア碁大会_対局者一覧_保存第３局.xlsx'
    save_execel_file_name = 'ペア碁大会_対局者一覧.xlsx'
    # 'result'もしくは対戦番号(1～)
    cmd = 'result'

    if cmd == 'result':
        write_result(excel_file_name, save_execel_file_name)
    else:
        taisenNo = int(cmd)
        player_decision(taisenNo, excel_file_name, save_execel_file_name)


