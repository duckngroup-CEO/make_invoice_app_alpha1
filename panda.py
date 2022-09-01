# 日付
import shutil
import pandas as pd
import os

# =======================================================
# パス定義
ORIGINAL_FILEPATH = "original_files/original_data.xlsx"
TEMPORARY_FILEPATH = "temporary_files/temporary_data.xlsx"
OUTPUT_FILEPATH = "output_files/output_data/output_data.xlsx"
OUTPUT_INVOICES_DIRPATH = "output_files/output_invoices/"
INVOICE_TEMPLATES = "static/invoice_templates/invoice_template1.xlsx"
# =======================================================
# メインの関数の定義
def main():
    # データ削除実行
    delete_data_by_index(11)

    # データ追加実行
    data_list = ["2021/4/30(追加)", "埼玉商事 株式会社", "海老コロッケ", 1200, 180, 10000]
    create_data(data_list)

    # データ編集実行
    index=3
    data_list=["2021-04-30 00:00:00(編集)", "埼玉商事 株式会社", "海老コロッケ", 1000, 1000, 1000]
    update_data(index,data_list)

    # 完成データの吐き出し
    shutil.copyfile(TEMPORARY_FILEPATH,OUTPUT_FILEPATH)

    # 請求書への埋め込み
    # 未完成ここを完成させる！！！！！！！！！！！！！！！！！！！！！！！！
    make_invoice(OUTPUT_INVOICES_DIRPATH)

# =======================================================
# 共通で使う関数の定義
# エクセルファイルの読み込み
def read_excel(filepath):
    df = pd.read_excel(filepath, index_col=0,dtype={"売上日":str})
    return df

# エクセルファイルの吐き出し
def output_excel(df, filepath):
    df.to_excel(filepath, sheet_name="売上", index=True, header=True)
    return df

# ファイル存在チェック関数
def file_exist_check(filepath):
    is_file = os.path.isfile(filepath)
    if is_file:
        return is_file
    else:
        pass

# 現在存在するインデックスを調べる関数
def make_index_list(df):
    index_list = list(df.index)
    return index_list
# =======================================================
# 機能関数
# 削除関数(特定のindex番号のデータだけ消去）
def delete_data_by_index(index):
    if file_exist_check(TEMPORARY_FILEPATH):
        df_tem = read_excel(TEMPORARY_FILEPATH)
        if index in make_index_list(df_tem):
            df_tem = df_tem.drop([index])
            df_tem = output_excel(df_tem, TEMPORARY_FILEPATH)
            return df_tem
        else:
            print("インデックスがありません")
    else:
        df = read_excel(ORIGINAL_FILEPATH)
        df_tem = output_excel(df, TEMPORARY_FILEPATH)
        if index in make_index_list(df_tem):
            df_tem = df_tem.drop([index])
            df_tem = output_excel(df_tem, TEMPORARY_FILEPATH)
            return df_tem
        else:
            print("インデックスがありません")
# =======================================================
# データ追加関数CREATE
def create_data(data_list):
    if file_exist_check(TEMPORARY_FILEPATH):
        df_tem = read_excel(TEMPORARY_FILEPATH)
        last_index = max(make_index_list(df_tem)) + 1
        add_df = pd.DataFrame([data_list,], columns = ["売上日", "顧客名", "商品名", "単価", "数量", "合計"], index=[last_index,])
        df = pd.concat([df_tem, add_df])
        df = output_excel(df, TEMPORARY_FILEPATH)
        return df
    else:
        df = read_excel(ORIGINAL_FILEPATH)
        df_tem = output_excel(df, TEMPORARY_FILEPATH)
        last_index = max(make_index_list(df_tem)) + 1
        add_df = pd.DataFrame([data_list,], columns = ["売上日", "顧客名", "商品名", "単価", "数量", "合計"], index=[last_index,])
        df = pd.concat([df_tem, add_df])
        df = output_excel(df, TEMPORARY_FILEPATH)
        return df
# =======================================================
# データ更新関数
# UPDATE
def update_data(index,data_list):
    if file_exist_check(TEMPORARY_FILEPATH):
        df_tem = read_excel(TEMPORARY_FILEPATH)
        if index in make_index_list(df_tem):
            change_df = pd.DataFrame([data_list,], columns = ["売上日", "顧客名", "商品名", "単価", "数量", "合計"], index=[index,])
            df_tem.update(change_df)
            df = output_excel(df_tem, TEMPORARY_FILEPATH)
            return df
        else:
            print("更新するindexがありません")
            pass

    else:
        df = read_excel(ORIGINAL_FILEPATH)
        df_tem = output_excel(df, TEMPORARY_FILEPATH)
        if index in make_index_list(df_tem):
            #updateする
                change_df = pd.DataFrame([data_list,], columns = ["売上日", "顧客名", "商品名", "単価", "数量", "合計"], index=[index,])
                df_tem.update(change_df)
                df = output_excel(df_tem, TEMPORARY_FILEPATH)
                return df
        else:
            print("更新するindexがありません")
            pass

# 請求書作成
def make_invoice(filepath):
    # ここを作り込む！！！！！！！！！！！！！！！！！！！！！！！！！
    pass
# =======================================================
# メイン関数の実行
if __name__ == "__main__":
    main()