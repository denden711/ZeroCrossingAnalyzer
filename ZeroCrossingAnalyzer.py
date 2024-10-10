import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime

# 間隔の設定（例: 0.00002秒）
TIME_INTERVAL_THRESHOLD = 0.00002

def find_zero_crossing_times(file_path):
    """CSVファイルから電圧が0を通過する時刻、周期、方向を計算し、指定された時間間隔が空いた場合に記録"""
    try:
        # CSVファイルをエンコーディングと高精度オプションを指定して読み込む
        df = pd.read_csv(file_path, encoding='shift_jis', float_precision='high')

        # 左から4列目が時間、5列目が電圧
        time_column = df.iloc[:, 3]
        voltage_column = df.iloc[:, 4]

        zero_crossing_times = []
        cycle_count = 0  # 周期カウンタ
        last_recorded_time = -float('inf')  # 最後に記録したゼロ交差時刻

        # データが空の場合のエラーハンドリング
        if len(time_column) == 0 or len(voltage_column) == 0:
            raise ValueError(f"{file_path} のデータが空です。")

        # 電圧が0を通過する瞬間を探す
        for i in range(1, len(voltage_column)):
            # ゼロ交差が発生した場合
            if (voltage_column[i-1] < 0 and voltage_column[i] >= 0) or (voltage_column[i-1] > 0 and voltage_column[i] <= 0):
                # 直前の時刻と現在の時刻を線形補間して、正確な交差時刻を求める
                zero_crossing_time = time_column[i-1] + (time_column[i] - time_column[i-1]) * (0 - voltage_column[i-1]) / (voltage_column[i] - voltage_column[i-1])

                # 最後に記録した時刻から指定時間間隔が経過しているか確認
                if zero_crossing_time - last_recorded_time >= TIME_INTERVAL_THRESHOLD:
                    # 上昇か下降かを判定
                    direction = "上昇（プラスに向かう）" if voltage_column[i-1] < 0 and voltage_column[i] >= 0 else "下降（マイナスに向かう）"

                    # 各交差点に対して周期数を記録
                    zero_crossing_times.append({
                        "ファイル名": os.path.basename(file_path),  # ファイル名を追加
                        "時刻": zero_crossing_time,
                        "周期": cycle_count,
                        "方向": direction
                    })

                    # 最後に記録した時刻を更新
                    last_recorded_time = zero_crossing_time

                    # 交差が一つの周期を示すので、上昇と下降を1セットとするなら、1周期進む
                    if direction == "上昇（プラスに向かう）":
                        cycle_count += 1

        return pd.DataFrame(zero_crossing_times)  # DataFrame形式で返す
    except UnicodeDecodeError:
        raise UnicodeDecodeError(f"{file_path} はエンコーディングに問題があります。shift_jis 以外のエンコーディングで保存されている可能性があります。")
    except pd.errors.EmptyDataError:
        raise ValueError(f"{file_path} は空のファイルです。")
    except Exception as e:
        raise RuntimeError(f"{file_path} の処理中にエラーが発生しました: {e}")

def select_csv_files():
    """ユーザーにCSVファイルを選択させるダイアログを表示"""
    root = tk.Tk()
    root.withdraw()  # メインウィンドウを隠す
    file_paths = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    return file_paths

def process_files(file_paths, output_excel_path):
    """選択されたCSVファイルを処理し、結果をExcelファイルに出力する"""
    with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
        for file_path in file_paths:
            try:
                df_zero_crossings = find_zero_crossing_times(file_path)
                if not df_zero_crossings.empty:
                    sheet_name = os.path.basename(file_path).replace('.csv', '')
                    # 出力時の精度を制御
                    df_zero_crossings.to_excel(writer, sheet_name=sheet_name, index=False, float_format="%.15f")
                    print(f"{file_path} の結果をExcelに出力しました。")
                else:
                    print(f"{file_path} にはゼロ交差点が見つかりませんでした。")
            except Exception as e:
                print(f"{file_path} の処理中にエラーが発生しました: {e}")

def main():
    """メインの処理を行う"""
    # CSVファイルを選択
    file_paths = select_csv_files()

    if file_paths:
        # 最初のファイル名を取得
        first_file_name = os.path.basename(file_paths[0]).replace('.csv', '')

        # 現在の時刻を取得してファイル名に組み込む
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 出力ファイル名を構築
        if len(file_paths) == 1:
            output_file_name = f"zero_crossing_results_{first_file_name}_{current_time}.xlsx"
        else:
            output_file_name = f"zero_crossing_results_{first_file_name}_and_others_{current_time}.xlsx"

        # 出力するExcelファイルのパスを指定（プログラムと同じフォルダ）
        output_excel_path = os.path.join(os.getcwd(), output_file_name)

        # ファイルを処理し、結果をExcelに出力
        process_files(file_paths, output_excel_path)
        print(f"結果は {output_excel_path} に保存されました。")
    else:
        print("CSVファイルが選択されませんでした")

if __name__ == "__main__":
    main()
