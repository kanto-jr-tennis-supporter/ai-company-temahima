# -*- coding: utf-8 -*-
"""他社の完成ファイルから会社データを全部抜いて「白紙テンプレート」を作る。

使い方:
    python3 strip_to_blank.py <入力ファイル> <出力ファイル> <抜く会社のプロフィールJSON>

やること:
  1. 全シート・全セルを走査
  2. 「様式のラベル」(記載要領・見出し・業種名など固定文言)は残す
  3. それ以外の入力値（会社データ・数値・チェック印）を全部クリア
  4. 残ったセルを一覧表示して、消し漏れがないか目視できるようにする

判定ルール:
  - 保護リスト(protect)に載っている文字列を含むセル → 残す
  - 20文字以上の日本語文（記載要領など） → 残す
  - それ以外で値が入っているセル → クリア
"""
import sys, json, re, shutil

def is_label(v):
    """様式に元から印字されているラベルかどうか"""
    if not isinstance(v, str):
        return False
    s = v.strip()
    if len(s) >= 20:              # 記載要領・注意書きなど長文
        return True
    # 見出し・項目名によく出る語
    label_words = [
        '※', '年', '月', '日', '円', '人', '点', '欄', '記載', '様式', '番号', '氏名',
        '住所', '名称', '商号', '電話', 'ＦＡＸ', 'FAX', 'フリガナ', 'ﾌ', 'ﾘ', 'ｶﾞ', 'ﾅ',
        '合計', '計', '申請', '営業', '工事', '建設', '資格', 'competition',
        '一式', '大工', '左官', 'とび', '石', '屋根', '電気', '管', 'タイル', '鋼',
        '鉄筋', '舗装', 'しゅん', '板金', 'ガラス', '塗装', '防水', '内装', '機械',
        '熱絶縁', '通信', '造園', 'さく井', '建具', '水道', '消防', '清掃', '解体',
        '新規', '更新', '外資', '国名', '比率', '職員', '設立', 'みなし', '大企業',
        '代表', '担当', '郵便', '所在', '区域', '業種', '区分', '評点', '完成',
        '北海道', '東北', '関東', '東海', '北陸', '近畿', '中国', '四国', '九州', '沖縄',
        '全国', '殿', '省', '局', '課', '長', '該当', '組合', '証明', '希望', '部局',
    ]
    return any(w in s for w in label_words)

def main():
    src, dst, profile_path = sys.argv[1], sys.argv[2], sys.argv[3]
    profile = json.load(open(profile_path))
    needles = [str(x) for x in profile['needles'] if str(x).strip()]

    shutil.copy(src, dst)

    if dst.endswith('.xls'):
        import xlrd
        from xlutils.copy import copy as xlcopy
        rb = xlrd.open_workbook(src, formatting_info=True)
        wb = xlcopy(rb)
        cleared, kept = [], []
        for si, rs in enumerate(rb.sheets()):
            ws = wb.get_sheet(si)
            for r in range(rs.nrows):
                for c in range(rs.ncols):
                    v = rs.cell_value(r, c)
                    if v in (None, ''):
                        continue
                    sv = str(v)
                    hit = any(n in sv for n in needles)
                    if hit or (not is_label(v)):
                        ws.write(r, c, '')
                        cleared.append((rs.name, r, c, sv))
                    else:
                        kept.append((rs.name, r, c, sv))
        wb.save(dst)
    else:
        import openpyxl, warnings
        warnings.filterwarnings('ignore')
        keep_vba = dst.endswith('.xlsm')
        wb = openpyxl.load_workbook(dst, keep_vba=keep_vba)
        cleared, kept = [], []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None or str(v).strip() == '':
                        continue
                    sv = str(v)
                    if sv.startswith('='):     # 数式は残す
                        kept.append((ws.title, cell.coordinate, sv))
                        continue
                    hit = any(n in sv for n in needles)
                    if hit or (not is_label(v)):
                        try:
                            cell.value = None
                            cleared.append((ws.title, cell.coordinate, sv))
                        except AttributeError:
                            pass               # 結合セルの非アンカー
                    else:
                        kept.append((ws.title, cell.coordinate, sv))
        wb.save(dst)

    print(f"■ クリアしたセル: {len(cleared)}件")
    for x in cleared:
        print("   ", x)
    print(f"\n■ 残したセル（様式ラベル判定）: {len(kept)}件")
    print(f"\n保存: {dst}")

if __name__ == '__main__':
    main()
