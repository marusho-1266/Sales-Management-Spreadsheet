"""
CSV保存モジュール
スクレイピング結果をCSVファイルとして保存する
"""
import logging
import pandas as pd
from pathlib import Path
from .config import DATA_DIR

# ロガーを設定
logger = logging.getLogger(__name__)


def save_result_csv(df: pd.DataFrame, filename: str = 'upload_data.csv') -> Path:
    """
    スクレイピング結果をCSVファイルとして保存する
    
    Args:
        df: スクレイピング結果のDataFrame
        filename: 保存するファイル名（デフォルト: upload_data.csv）
    
    Returns:
        Path: 保存されたCSVファイルのパス
    
    Raises:
        Exception: CSVファイルの保存に失敗した場合
    """
    data_dir = Path(DATA_DIR)
    csv_path = data_dir / filename
    
    # 親ディレクトリが存在することを確認（存在しない場合は作成）
    data_dir.mkdir(parents=True, exist_ok=True)
    
    # UTF-8 BOM付きで保存（Excel互換性のため）
    try:
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    except Exception as e:
        error_message = f"CSVファイルの保存に失敗しました: {csv_path}, エラー: {e}"
        logger.error(error_message)
        raise Exception(error_message) from e
    
    print(f"CSVファイルを保存しました: {csv_path}")
    return csv_path
