import logging

# ログ設定を定義
def setup_logger(log_file="app.log"):
    logger = logging.getLogger("ExcelLogger")
    logger.setLevel(logging.DEBUG)

    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # ファイルハンドラ
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)

    # フォーマット
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # ハンドラを追加
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger

# グローバルで使用するロガー
logger = setup_logger()
