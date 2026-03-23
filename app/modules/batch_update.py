"""
FALCON2 - 一括更新モジュール
コマンドラインから条件式を使って牛のデータを一括更新
例: ENTR=2024-04-01 : ID>0
例: PEN=2 : LACT=2 DIM=100-200
"""

import re
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime
import logging

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine

logger = logging.getLogger(__name__)


class BatchUpdate:
    """一括更新クラス"""
    
    # cowテーブルの更新可能なカラム
    UPDATEABLE_COLUMNS = {
        'ENTR': 'entr',
        'PEN': 'pen',
        'BRD': 'brd',
        'BTHD': 'bthd',
        'FRM': 'frm',
        'LACT': 'lact',
        'RC': 'rc',
    }
    
    def __init__(self, db_handler: DBHandler, formula_engine: FormulaEngine):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
        """
        self.db = db_handler
        self.formula_engine = formula_engine
    
    def parse_command(self, command: str) -> Tuple[Dict[str, str], str]:
        """
        コマンドをパース
        
        Args:
            command: コマンド文字列（例: "PEN=100：LACT=0" または "ENTR=2024-04-01 : ID>0"）
        
        Returns:
            (更新項目辞書, 条件式文字列)
            例: ({'PEN': '100'}, 'LACT=0')
        """
        # 全角コロンも半角に統一して分割（左側が更新項目、右側が条件）
        normalized = command.replace('：', ':')
        parts = normalized.split(':', 1)
        
        if len(parts) != 2:
            raise ValueError("コマンドは '項目=値：条件' の形式で指定してください")
        
        update_part = parts[0].strip()
        condition_part = parts[1].strip() if len(parts) > 1 else ""
        
        # 更新項目をパース（例: "PEN=100"）
        update_items = {}
        if '=' in update_part:
            key, value = update_part.split('=', 1)
            key = key.strip().upper()
            value = value.strip()
            update_items[key] = value
        else:
            raise ValueError("更新項目は '項目=値' の形式で指定してください")
        
        return update_items, condition_part
    
    def evaluate_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """
        条件式を評価
        
        サポートする条件:
        - ID>0, ID<100, ID=1234
        - LACT=2, LACT>1
        - DIM=100-200 (範囲指定)
        - 複数条件はスペース区切りでAND結合
        
        Args:
            cow: 牛データ
            condition: 条件式文字列
        
        Returns:
            条件を満たす場合はTrue
        """
        if not condition:
            return True
        
        # 条件を分割（スペース区切りでAND結合）
        conditions = condition.split()
        
        for cond in conditions:
            if not self._evaluate_single_condition(cow, cond):
                return False
        
        return True
    
    def _evaluate_single_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """
        単一条件を評価
        
        Args:
            cow: 牛データ
            condition: 単一条件式（例: "ID>0", "LACT=2", "DIM=100-200"）
        
        Returns:
            条件を満たす場合はTrue
        """
        condition_upper = condition.upper()
        
        # ID条件（大文字小文字を区別しない）
        if condition_upper.startswith('ID'):
            return self._evaluate_id_condition(cow, condition)
        
        # LACT条件（大文字小文字を区別しない）
        if condition_upper.startswith('LACT'):
            return self._evaluate_lact_condition(cow, condition)
        
        # DIM条件（範囲指定対応、大文字小文字を区別しない）
        if condition_upper.startswith('DIM'):
            return self._evaluate_dim_condition(cow, condition)
        
        # PEN条件（例: PEN=100）
        if condition_upper.startswith('PEN'):
            return self._evaluate_pen_condition(cow, condition)
        
        # その他の数値比較（将来拡張用）
        if '=' in condition or '>' in condition or '<' in condition:
            return self._evaluate_generic_condition(cow, condition)
        
        return True
    
    def _evaluate_id_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """ID条件を評価（例: ID>0, ID=1234）"""
        cow_id = cow.get('cow_id', '')
        try:
            cow_id_int = int(cow_id) if cow_id.isdigit() else 0
        except:
            cow_id_int = 0
        
        if '>=' in condition:
            threshold = int(condition.split('>=')[1])
            return cow_id_int >= threshold
        elif '<=' in condition:
            threshold = int(condition.split('<=')[1])
            return cow_id_int <= threshold
        elif '>' in condition:
            threshold = int(condition.split('>')[1])
            return cow_id_int > threshold
        elif '<' in condition:
            threshold = int(condition.split('<')[1])
            return cow_id_int < threshold
        elif '=' in condition:
            threshold = int(condition.split('=')[1])
            return cow_id_int == threshold
        
        return True
    
    def _evaluate_lact_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """LACT条件を評価（例: LACT=2, LACT>1）"""
        lact = cow.get('lact') or 0
        
        if '>=' in condition:
            threshold = int(condition.split('>=')[1])
            return lact >= threshold
        elif '<=' in condition:
            threshold = int(condition.split('<=')[1])
            return lact <= threshold
        elif '>' in condition:
            threshold = int(condition.split('>')[1])
            return lact > threshold
        elif '<' in condition:
            threshold = int(condition.split('<')[1])
            return lact < threshold
        elif '=' in condition:
            threshold = int(condition.split('=')[1])
            return lact == threshold
        
        return True
    
    def _evaluate_dim_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """DIM条件を評価（例: DIM=100-200, DIM>100）"""
        # DIMは計算項目なので、FormulaEngineで計算
        cow_auto_id = cow.get('auto_id')
        if not cow_auto_id:
            return False
        
        try:
            calculated = self.formula_engine.calculate(cow_auto_id, 'DIM')
            dim = calculated.get('DIM')
            if dim is None:
                return False
            dim = int(dim)  # 整数に変換
        except Exception as e:
            logger.debug(f"DIM計算エラー: cow_auto_id={cow_auto_id}, エラー={e}")
            return False
        
        # 範囲指定（例: DIM=100-200）
        if '=' in condition and '-' in condition:
            range_part = condition.split('=')[1]
            if '-' in range_part:
                min_dim, max_dim = map(int, range_part.split('-'))
                return min_dim <= dim <= max_dim
        
        # 通常の比較
        if '>=' in condition:
            threshold = int(condition.split('>=')[1])
            return dim >= threshold
        elif '<=' in condition:
            threshold = int(condition.split('<=')[1])
            return dim <= threshold
        elif '>' in condition:
            threshold = int(condition.split('>')[1])
            return dim > threshold
        elif '<' in condition:
            threshold = int(condition.split('<')[1])
            return dim < threshold
        elif '=' in condition:
            threshold = int(condition.split('=')[1])
            return dim == threshold
        
        return True
    
    def _evaluate_pen_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """PEN条件を評価（例: PEN=100）"""
        pen = cow.get('pen')
        if pen is None or str(pen).strip() == '':
            return False
        try:
            pen_val = int(pen)
        except (ValueError, TypeError):
            return False
        if '>=' in condition:
            threshold = int(condition.split('>=')[1])
            return pen_val >= threshold
        elif '<=' in condition:
            threshold = int(condition.split('<=')[1])
            return pen_val <= threshold
        elif '>' in condition:
            threshold = int(condition.split('>')[1])
            return pen_val > threshold
        elif '<' in condition:
            threshold = int(condition.split('<')[1])
            return pen_val < threshold
        elif '=' in condition:
            threshold = int(condition.split('=')[1])
            return pen_val == threshold
        return True
    
    def _evaluate_generic_condition(self, cow: Dict[str, Any], condition: str) -> bool:
        """汎用条件を評価（将来拡張用）"""
        # 現時点では基本的な実装のみ
        return True
    
    def get_matching_cows(self, condition: str) -> List[Dict[str, Any]]:
        """
        条件に合致する牛のリストを返す（更新は行わない）。
        
        Args:
            condition: 条件式文字列（例: 'LACT=0'）
        
        Returns:
            条件に合致する牛の辞書のリスト
        """
        all_cows = self.db.get_all_cows()
        matching = []
        for cow in all_cows:
            if self.evaluate_condition(cow, condition):
                matching.append(cow)
        return matching
    
    def execute_update(self, update_items: Dict[str, str], condition: str) -> int:
        """
        一括更新を実行
        
        Args:
            update_items: 更新項目辞書（例: {'ENTR': '2024-04-01'}）
            condition: 条件式文字列（例: 'ID>0'）
        
        Returns:
            更新された牛の数
        """
        # 全牛を取得
        all_cows = self.db.get_all_cows()
        
        # 条件に合致する牛を抽出
        matching_cows = []
        for cow in all_cows:
            if self.evaluate_condition(cow, condition):
                matching_cows.append(cow)
        
        if not matching_cows:
            print(f"条件に合致する牛が見つかりませんでした: {condition}")
            logger.info(f"条件に合致する牛が見つかりませんでした: {condition}")
            return 0
        
        # 更新を実行
        updated_count = 0
        print(f"条件に合致する牛: {len(matching_cows)}頭")
        
        for cow in matching_cows:
            cow_auto_id = cow.get('auto_id')
            if not cow_auto_id:
                continue
            
            # 更新データを構築
            update_data = {}
            for key, value in update_items.items():
                if key in self.UPDATEABLE_COLUMNS:
                    db_column = self.UPDATEABLE_COLUMNS[key]
                    update_data[db_column] = value
                else:
                    print(f"警告: 更新できない項目が指定されました: {key}")
            
            if update_data:
                try:
                    self.db.update_cow(cow_auto_id, update_data)
                    updated_count += 1
                    print(f"更新: cow_id={cow.get('cow_id')}, 更新内容={update_data}")
                    logger.info(f"牛を更新: cow_id={cow.get('cow_id')}, auto_id={cow_auto_id}, 更新内容={update_data}")
                except Exception as e:
                    print(f"エラー: cow_id={cow.get('cow_id')} の更新に失敗: {e}")
                    logger.error(f"牛の更新に失敗: cow_id={cow.get('cow_id')}, auto_id={cow_auto_id}, エラー={e}")
            else:
                print(f"警告: cow_id={cow.get('cow_id')} の更新データが空です")
        
        print(f"一括更新が完了しました: {updated_count}頭を更新")
        logger.info(f"一括更新が完了しました: {updated_count}頭を更新")
        return updated_count

