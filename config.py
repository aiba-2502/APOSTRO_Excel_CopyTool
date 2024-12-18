# 設定値を大カテゴリに分けて辞書形式で定義
config = {
    # ファイルパス関連
    "files": {
        "value_file": r"C:\Users\aiba\Desktop\テスト仕様書\元\ClinicPOS試験_現金決済(ACE-100)_まとめて会計-v3.3.0.xlsx",  # 値をコピーする元のExcelファイルのパス
        "output_file": r"C:\Users\aiba\Desktop\テスト仕様書\ClinicKIOSK試験(バーコード)_現金(RT-300)決済(まとめて会計).xlsx",      # 出力先のExcelファイルのパス
    },

    # シート名関連
    "sheets": {
        "template_sheet_name": "現金会計_1人_決済中止",  # テンプレートとして使用するシート名（出力ファイル内）
        "value_sheet_name": "現金決済(ACE-100)_4人",        # 値をコピーする元のシート名（値ファイル内）
        "output_sheet_name": "OutputSheet",      # 複製したシートの名前（出力ファイル内での新しいシート名）
    },

    # コピー設定
    "copy_settings": {
        "copy_range": "C12:E19",  # コピーする範囲を指定（例: "A1:C10"）
        "paste_start": "B15",     # 出力ファイルで値を貼り付け始めるセルの位置（例: "A1"）
    },

    # 行の高さ調整関連
    "row_height_settings": {
        "default_font_size": 10,  # デフォルトフォントサイズ（行の高さ調整に使用）
        "min_row_height": 15,     # 最低行の高さ
        "line_height_multiplier": 2.0,  # 行の高さ調整の倍率
    }
}
