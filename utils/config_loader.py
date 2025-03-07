import json
import os

def load_config(config_path: str = "config.json") -> dict:
    """
    載入設定檔
    
    Args:
        config_path: 設定檔路徑，預設為 config.json
        
    Returns:
        設定檔內容的字典
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"找不到設定檔：{config_path}")
        
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f) 