"""
FALCON2 - Normalization Dictionary Generator
item_dictionary.json / event_dictionary.json から normalization 辞書を自動生成
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Set
import unicodedata


def normalize_japanese(text: str) -> str:
    """
    日本語テキストを正規化
    - 全角半角統一（全角に統一）
    - 記号除去
    - 空白正規化
    """
    if not text:
        return ""
    
    # 全角に統一
    text = unicodedata.normalize('NFKC', text)
    
    # 記号除去（カンマ、ピリオド、括弧など）
    text = re.sub(r'[，。、．（）()【】「」『』〈〉《》［］｛｝]', '', text)
    
    # 空白正規化（連続空白を1つに）
    text = re.sub(r'\s+', ' ', text)
    
    return text.strip()


def split_japanese_words(text: str) -> List[str]:
    """
    日本語テキストを分割可能な語に分割（助詞除去）
    
    例：
    "初回授精日数" → ["初回授精日数", "初回 授精 日数", "初回授精"]
    """
    if not text:
        return []
    
    results: Set[str] = set()
    results.add(text)  # 元の文字列
    
    # 数字や記号で分割
    parts = re.split(r'[0-9０-９・\-]', text)
    parts = [p.strip() for p in parts if p.strip()]
    
    if len(parts) > 1:
        # 分割された部分を結合
        results.add(' '.join(parts))
    
    # 2文字以上の部分文字列を生成（先頭から）
    if len(text) >= 2:
        for i in range(2, len(text) + 1):
            substring = text[:i]
            if len(substring) >= 2:
                results.add(substring)
    
    return sorted(list(results), key=len, reverse=True)


def generate_items_dict(item_dictionary: Dict) -> Dict[str, List[str]]:
    """
    item_dictionary から items.json を生成
    
    Args:
        item_dictionary: item_dictionary.json の内容
    
    Returns:
        items.json の内容
    """
    result: Dict[str, List[str]] = {}
    
    for item_key, item_data in item_dictionary.items():
        if not isinstance(item_data, dict):
            continue
        
        synonyms: Set[str] = set()
        
        # 1. item_key 自体（大文字・小文字）
        synonyms.add(item_key)
        synonyms.add(item_key.lower())
        
        # 2. display_name を取得
        display_name = item_data.get("display_name", "")
        if display_name:
            # 正規化版
            normalized = normalize_japanese(display_name)
            if normalized:
                synonyms.add(normalized)
            
            # 記号除去版
            no_symbol = re.sub(r'[^\w\s]', '', display_name)
            if no_symbol and no_symbol != display_name:
                synonyms.add(no_symbol.strip())
            
            # 分割可能な語
            split_words = split_japanese_words(display_name)
            for word in split_words:
                if word and len(word) >= 2:
                    synonyms.add(word)
        
        # 3. description から略称を抽出（オプション）
        # 注意：英語の説明文から単語を抽出しない
        description = item_data.get("description", "")
        if description:
            # 「略称: XXX」パターンを抽出（日本語説明内の略称のみ）
            match = re.search(r'略称[：:]\s*([A-Za-z0-9_]+)', description)
            if match:
                abbrev = match.group(1)
                # 略称が3文字以上で、大文字のみまたは大文字+数字の場合は追加
                if len(abbrev) >= 2 and (abbrev.isupper() or re.match(r'^[A-Z]+[0-9]*$', abbrev)):
                    synonyms.add(abbrev)
                    synonyms.add(abbrev.lower())
        
        # リストに変換してソート（重複除去済み）
        result[item_key] = sorted(list(synonyms))
    
    return result


def generate_events_dict(event_dictionary: Dict) -> Dict[str, List[str]]:
    """
    event_dictionary から events.json を生成
    
    Args:
        event_dictionary: event_dictionary.json の内容
    
    Returns:
        events.json の内容
    """
    result: Dict[str, List[str]] = {}
    
    for event_number_str, event_data in event_dictionary.items():
        if not isinstance(event_data, dict):
            continue
        
        # alias を event_key として使用
        alias = event_data.get("alias")
        if not alias:
            continue
        
        synonyms: Set[str] = set()
        
        # 1. alias 自体（大文字・小文字）
        synonyms.add(alias)
        synonyms.add(alias.lower())
        
        # 2. name_jp（日本語名）
        name_jp = event_data.get("name_jp", "")
        if name_jp:
            # 正規化版
            normalized = normalize_japanese(name_jp)
            if normalized:
                synonyms.add(normalized)
            
            # 記号除去版
            no_symbol = re.sub(r'[^\w\s]', '', name_jp)
            if no_symbol and no_symbol != name_jp:
                synonyms.add(no_symbol.strip())
            
            # 分割可能な語
            split_words = split_japanese_words(name_jp)
            for word in split_words:
                if word and len(word) >= 2:
                    synonyms.add(word)
        
        # 3. CALVING などの特殊マッピング
        if alias == "CALV":
            # CALVING としても登録
            if "CALVING" not in result:
                result["CALVING"] = []
            result["CALVING"].extend(list(synonyms))
        
        # 4. 妊娠鑑定のマッピング
        if alias == "PDP":
            if "PREG_POS" not in result:
                result["PREG_POS"] = []
            result["PREG_POS"].extend(list(synonyms))
        elif alias == "PDN":
            if "PREG_NEG" not in result:
                result["PREG_NEG"] = []
            result["PREG_NEG"].extend(list(synonyms))
        
        # 通常の alias をキーとして登録
        if alias not in result:
            result[alias] = []
        result[alias].extend(list(synonyms))
    
    # 重複除去とソート
    for key in result:
        result[key] = sorted(list(set(result[key])))
    
    return result


def load_terms_dict(terms_path: Path) -> Dict[str, List[str]]:
    """
    terms.json を読み込む（手動定義・固定）
    
    Args:
        terms_path: terms.json のパス
    
    Returns:
        terms.json の内容
    """
    if terms_path.exists():
        try:
            with open(terms_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"terms.json読み込みエラー: {e}")
    
    # デフォルト
    return {
        "AVG": ["平均"],
        "COUNT": ["頭数", "件数"],
        "LIST": ["教えて", "一覧"],
        "SCATTER": ["散布図"],
        "TABLE": ["表"]
    }


def main():
    """メイン処理"""
    # パス設定
    base_dir = Path(__file__).parent.parent
    config_default_dir = base_dir / "config_default"
    normalization_dir = base_dir / "normalization"
    
    # ディレクトリ作成
    normalization_dir.mkdir(exist_ok=True)
    
    # 入力ファイル
    item_dict_path = config_default_dir / "item_dictionary.json"
    event_dict_path = config_default_dir / "event_dictionary.json"
    
    # 出力ファイル
    items_output_path = normalization_dir / "items.json"
    events_output_path = normalization_dir / "events.json"
    terms_output_path = normalization_dir / "terms.json"
    
    # item_dictionary 読み込み
    if not item_dict_path.exists():
        print(f"エラー: {item_dict_path} が見つかりません")
        return
    
    with open(item_dict_path, 'r', encoding='utf-8') as f:
        item_dictionary = json.load(f)
    
    # event_dictionary 読み込み
    if not event_dict_path.exists():
        print(f"エラー: {event_dict_path} が見つかりません")
        return
    
    with open(event_dict_path, 'r', encoding='utf-8') as f:
        event_dictionary = json.load(f)
    
    # 辞書生成
    print("items.json を生成中...")
    items_dict = generate_items_dict(item_dictionary)
    
    print("events.json を生成中...")
    events_dict = generate_events_dict(event_dictionary)
    
    # 保存
    with open(items_output_path, 'w', encoding='utf-8') as f:
        json.dump(items_dict, f, ensure_ascii=False, indent=2)
    print(f"[OK] {items_output_path} を保存しました")
    
    with open(events_output_path, 'w', encoding='utf-8') as f:
        json.dump(events_dict, f, ensure_ascii=False, indent=2)
    print(f"[OK] {events_output_path} を保存しました")
    
    # terms.json は既存があれば保持、なければデフォルトを保存
    terms_dict = load_terms_dict(terms_output_path)
    if not terms_output_path.exists():
        with open(terms_output_path, 'w', encoding='utf-8') as f:
            json.dump(terms_dict, f, ensure_ascii=False, indent=2)
        print(f"[OK] {terms_output_path} を保存しました（デフォルト）")
    else:
        print(f"[OK] {terms_output_path} は既存のものを保持しました")
    
    print("\n完了: normalization 辞書を生成しました")


if __name__ == "__main__":
    main()

