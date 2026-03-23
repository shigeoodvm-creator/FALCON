"""
FALCON2 - QueryRouter V2（改修版）
日本語・改行コマンドからExecutionPlanを生成するエンジン
"""

import re
import json
import logging
from typing import Dict, Any, Optional, List, Tuple
from pathlib import Path
from datetime import datetime, date

from modules.execution_plan import (
    ExecutionPlan, AnalysisType, GraphType, Condition, GroupBy
)

logger = logging.getLogger(__name__)


class QueryRouterV2:
    """
    日本語・改行コマンドからExecutionPlanを生成
    
    機能:
    - 複数行コマンドの解析
    - 辞書解決（item_dictionary.json, event_dictionary.json）
    - 正規化（全角記号、カンマ等）
    - 条件行のパース
    - DIM区分の処理
    - グラフ種類の処理
    - イベントグルーピング
    """
    
    # イベントグルーピング辞書
    EVENT_GROUPS = {
        "授精": [200, 201],  # AI + ET
        "分娩": [202],
        "妊娠鑑定プラス": [303, 304],  # PDP + PDP2
        "妊娠鑑定マイナス": [302],  # PDN
        "PAGプラス": [307],  # PAGP
        "PAGマイナス": [306],  # PAGN
    }
    
    # キーワード条件マッピング
    KEYWORD_CONDITIONS = {
        "妊娠牛": {"item_key": "RC", "operator": "=", "value": 5},  # RC_PREGNANT
        "空胎": {"item_key": "RC", "operator": "=", "value": 4},  # RC_OPEN
        "乾乳牛": {"item_key": "RC", "operator": "=", "value": 6},  # RC_DRY
        "未授精": {"item_key": "BRED", "operator": "=", "value": 0},
        "初回授精": {"item_key": "BRED", "operator": "=", "value": 1},
    }
    
    def __init__(self, item_dictionary_path: Optional[Path] = None,
                 event_dictionary_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            item_dictionary_path: item_dictionary.json のパス
            event_dictionary_path: event_dictionary.json のパス
        """
        self.item_dictionary: Dict[str, Any] = {}
        self.event_dictionary: Dict[str, Any] = {}
        self.item_key_by_display_name: Dict[str, str] = {}  # display_name -> item_key
        self.item_key_by_key: Dict[str, str] = {}  # key -> item_key（重複チェック用）
        self.event_key_by_name_jp: Dict[str, str] = {}  # name_jp -> event_key
        self.event_key_by_alias: Dict[str, str] = {}  # alias -> event_key
        
        self._load_dictionaries(item_dictionary_path, event_dictionary_path)
    
    def _load_dictionaries(self, item_dict_path: Optional[Path],
                          event_dict_path: Optional[Path]):
        """辞書を読み込む"""
        # item_dictionary.json
        if item_dict_path and item_dict_path.exists():
            try:
                with open(item_dict_path, 'r', encoding='utf-8') as f:
                    self.item_dictionary = json.load(f)
                
                # インデックスを作成
                for item_key, item_def in self.item_dictionary.items():
                    if isinstance(item_def, dict):
                        display_name = item_def.get("display_name", "")
                        if display_name:
                            # display_name -> item_key のマッピング
                            if display_name not in self.item_key_by_display_name:
                                self.item_key_by_display_name[display_name] = item_key
                            else:
                                # 重複がある場合は警告
                                logger.warning(f"display_name重複: {display_name} -> {item_key} (既存: {self.item_key_by_display_name[display_name]})")
                        
                        # key（item_key自体）のマッピング
                        self.item_key_by_key[item_key] = item_key
                        
                        # descriptionから略称を抽出（例: "略称: Days In Milk" -> "DIM"）
                        description = item_def.get("description", "")
                        if description:
                            # "略称: XXX" パターンを検索
                            match = re.search(r'略称[：:]\s*([A-Z0-9_]+)', description)
                            if match:
                                abbrev = match.group(1)
                                if abbrev not in self.item_key_by_key:
                                    self.item_key_by_key[abbrev] = item_key
            except Exception as e:
                logger.error(f"item_dictionary.json読み込みエラー: {e}")
        
        # event_dictionary.json
        if event_dict_path and event_dict_path.exists():
            try:
                with open(event_dict_path, 'r', encoding='utf-8') as f:
                    self.event_dictionary = json.load(f)
                
                # インデックスを作成
                for event_number_str, event_data in self.event_dictionary.items():
                    if isinstance(event_data, dict):
                        name_jp = event_data.get("name_jp", "")
                        alias = event_data.get("alias", "")
                        
                        if name_jp:
                            self.event_key_by_name_jp[name_jp] = event_number_str
                        if alias:
                            self.event_key_by_alias[alias] = event_number_str
            except Exception as e:
                logger.error(f"event_dictionary.json読み込みエラー: {e}")
    
    def parse(self, analysis_type: str, period_start: Optional[str],
              period_end: Optional[str], multiline_text: str) -> ExecutionPlan:
        """
        複数行コマンドを解析してExecutionPlanを生成
        
        Args:
            analysis_type: 分析種別（"list", "agg", "eventcount", "graph", "repro"）
            period_start: 期間開始（YYYY-MM-DD）
            period_end: 期間終了（YYYY-MM-DD）
            multiline_text: 複数行コマンドテキスト
        
        Returns:
            ExecutionPlan
        """
        # 分析種別を安全に変換
        try:
            if isinstance(analysis_type, AnalysisType):
                analysis_type_enum = analysis_type
            else:
                analysis_type_enum = AnalysisType(analysis_type)
        except (ValueError, KeyError):
            # 無効な分析種別の場合はエラーを返す
            plan = ExecutionPlan(
                analysis_type=AnalysisType.LIST,  # デフォルト
                period_start=period_start,
                period_end=period_end
            )
            plan.errors.append(f"無効な分析種別: {analysis_type}")
            return plan
        
        plan = ExecutionPlan(
            analysis_type=analysis_type_enum,
            period_start=period_start,
            period_end=period_end
        )
        
        # 正規化：全角記号を半角に、空行を除去
        lines = self._normalize_lines(multiline_text)
        
        if not lines:
            plan.errors.append("コマンドが空です")
            return plan
        
        # 分析種別ごとに解析
        if analysis_type == "list":
            self._parse_list(lines, plan)
        elif analysis_type == "agg":
            self._parse_agg(lines, plan)
        elif analysis_type == "eventcount":
            self._parse_eventcount(lines, plan)
        elif analysis_type == "graph":
            self._parse_graph(lines, plan)
        elif analysis_type == "repro":
            self._parse_repro(lines, plan)
        else:
            plan.errors.append(f"未対応の分析種別: {analysis_type}")
        
        return plan
    
    def _normalize_lines(self, text: str) -> List[str]:
        """
        テキストを正規化して行リストに変換
        
        - 全角記号を半角に（、→,、＞＜＝→><=）
        - 空行を除去
        """
        lines = []
        for line in text.split('\n'):
            # 全角記号を半角に変換
            line = line.replace('、', ',')
            line = line.replace('＞', '>')
            line = line.replace('＜', '<')
            line = line.replace('＝', '=')
            line = line.replace('：', ':')
            
            # 前後の空白を削除
            line = line.strip()
            
            # 空行は無視
            if line:
                lines.append(line)
        
        return lines
    
    def _parse_list(self, lines: List[str], plan: ExecutionPlan):
        """リスト解析"""
        if not lines:
            plan.errors.append("1行目に表示項目を指定してください")
            return
        
        # 1行目: 表示項目（カンマ区切りまたはスペース区切り）
        first_line = lines[0]
        columns = self._parse_item_list(first_line)
        if not columns:
            plan.errors.append("1行目に表示項目を指定してください")
        else:
            plan.columns = columns
        
        # 2行目以降: 絞り込み条件
        for line in lines[1:]:
            condition = self._parse_condition(line)
            if condition:
                plan.conditions.append(condition)
            else:
                plan.warnings.append(f"条件として解釈できませんでした: {line}")
    
    def _parse_agg(self, lines: List[str], plan: ExecutionPlan):
        """集計解析"""
        if not lines:
            plan.errors.append("1行目に集計項目を指定してください")
            return
        
        # 1行目: 集計項目
        first_line = lines[0]
        aggregate_items = self._parse_item_list(first_line)
        if not aggregate_items:
            plan.errors.append("1行目に集計項目を指定してください")
        else:
            plan.aggregate_items = aggregate_items
        
        # 2行目: 分けて見る基準
        if len(lines) > 1:
            group_by_line = lines[1]
            group_by = self._parse_group_by(group_by_line)
            if group_by:
                plan.group_by = group_by
        
        # 3行目以降: 絞り込み条件
        for line in lines[2:]:
            condition = self._parse_condition(line)
            if condition:
                plan.conditions.append(condition)
    
    def _parse_eventcount(self, lines: List[str], plan: ExecutionPlan):
        """イベント集計解析"""
        if not lines:
            plan.errors.append("1行目に集計したいイベントを指定してください")
            return
        
        # 1行目: 集計したいイベント
        first_line = lines[0]
        event_keys = self._parse_event_list(first_line)
        if not event_keys:
            plan.errors.append("1行目に集計したいイベントを指定してください")
        else:
            plan.event_keys = event_keys
        
        # 2行目: 分けて見る基準
        if len(lines) > 1:
            group_by_line = lines[1]
            group_by = self._parse_group_by(group_by_line)
            if group_by:
                plan.group_by = group_by
        
        # 3行目以降: 絞り込み条件
        for line in lines[2:]:
            condition = self._parse_condition(line)
            if condition:
                plan.conditions.append(condition)
    
    def _parse_graph(self, lines: List[str], plan: ExecutionPlan):
        """グラフ解析"""
        if not lines:
            plan.errors.append("1行目に縦軸項目を指定してください")
            return
        
        # 1行目: 縦軸項目
        y_item = self._resolve_item(lines[0])
        if not y_item:
            plan.errors.append(f"1行目の縦軸項目を解決できませんでした: {lines[0]}")
        else:
            plan.y_axis_item = y_item
        
        # 2行目: 横軸項目
        if len(lines) > 1:
            x_item = self._resolve_item(lines[1])
            if x_item:
                plan.x_axis_item = x_item
        
        # 3行目: グラフ種類
        if len(lines) > 2:
            graph_type = self._parse_graph_type(lines[2])
            if graph_type:
                plan.graph_type = graph_type
        
        # 4行目: 色・線を分ける基準
        if len(lines) > 3:
            split_item = self._resolve_item(lines[3])
            if split_item:
                plan.split_by_item = split_item
        
        # 5行目以降: 絞り込み条件
        for line in lines[4:]:
            condition = self._parse_condition(line)
            if condition:
                plan.conditions.append(condition)
    
    def _parse_repro(self, lines: List[str], plan: ExecutionPlan):
        """繁殖分析解析"""
        if not lines:
            plan.errors.append("1行目に繁殖指標を指定してください")
            return
        
        # 1行目: 繁殖指標
        first_line = lines[0]
        repro_metric = self._resolve_repro_metric(first_line)
        if not repro_metric:
            plan.errors.append(f"1行目の繁殖指標を解決できませんでした: {first_line}")
        else:
            plan.repro_metric = repro_metric
        
        # 2行目: 分けて見る基準
        if len(lines) > 1:
            group_by_line = lines[1]
            group_by = self._parse_group_by(group_by_line)
            if group_by:
                plan.group_by = group_by
        
        # 3行目以降: 絞り込み条件
        for line in lines[2:]:
            condition = self._parse_condition(line)
            if condition:
                plan.conditions.append(condition)
    
    def _parse_item_list(self, text: str) -> List[str]:
        """項目リストを解析（カンマ区切りまたはスペース区切り）"""
        # カンマ区切りまたはスペース区切りで分割
        items = re.split(r'[,，\s]+', text)
        resolved_items = []
        
        for item in items:
            item = item.strip()
            if not item:
                continue
            
            resolved = self._resolve_item(item)
            if resolved:
                resolved_items.append(resolved)
        
        return resolved_items
    
    def _parse_event_list(self, text: str) -> List[str]:
        """イベントリストを解析"""
        # カンマ区切りまたはスペース区切りで分割
        items = re.split(r'[,，\s]+', text)
        resolved_events = []
        
        for item in items:
            item = item.strip()
            if not item:
                continue
            
            # イベントグルーピングを確認
            if item in self.EVENT_GROUPS:
                # グルーピング名をそのまま使用（Executorで展開）
                resolved_events.append(item)
            else:
                # 個別イベントを解決
                resolved = self._resolve_event(item)
                if resolved:
                    resolved_events.append(resolved)
        
        return resolved_events
    
    def _parse_group_by(self, text: str) -> Optional[GroupBy]:
        """GROUP BY句を解析"""
        text = text.strip()
        
        # DIM区分チェック（DIM7/DIM14/DIM21/DIM30/DIM50）
        dim_bin_match = re.match(r'DIM(\d+)', text, re.IGNORECASE)
        if dim_bin_match:
            width = int(dim_bin_match.group(1))
            return GroupBy(bin_base="DIM", bin_width=width)
        
        # 通常の項目
        item_key = self._resolve_item(text)
        if item_key:
            return GroupBy(item_key=item_key)
        
        return None
    
    def _parse_graph_type(self, text: str) -> Optional[GraphType]:
        """グラフ種類を解析"""
        text = text.strip().lower()
        
        if "折れ線" in text or "line" in text:
            return GraphType.LINE
        elif "棒" in text or "bar" in text:
            return GraphType.BAR
        elif "プロット" in text or "scatter" in text:
            return GraphType.SCATTER
        elif "箱ひげ" in text or "box" in text:
            return GraphType.BOX
        elif "生存曲線" in text or "survival" in text:
            return GraphType.SURVIVAL
        
        return None
    
    def _parse_condition(self, text: str) -> Optional[Condition]:
        """条件行を解析"""
        text = text.strip()
        
        # キーワード条件チェック
        for keyword, cond_dict in self.KEYWORD_CONDITIONS.items():
            if keyword in text:
                return Condition(
                    item_key=cond_dict["item_key"],
                    operator=cond_dict["operator"],
                    value=cond_dict["value"],
                    keyword=keyword
                )
        
        # 比較条件パース: <項目><演算子><値>
        # 例: "DIM>150", "産次>=2", "LACT=1"
        pattern = r'([^\s<>=!]+)\s*(>=|<=|!=|>|<|=)\s*(.+)'
        match = re.match(pattern, text)
        if match:
            item_str = match.group(1).strip()
            operator = match.group(2).strip()
            value_str = match.group(3).strip()
            
            # 項目を解決
            item_key = self._resolve_item(item_str)
            if not item_key:
                return None
            
            # 値を解析（数値または文字列）
            value = self._parse_value(value_str)
            
            return Condition(
                item_key=item_key,
                operator=operator,
                value=value
            )
        
        return None
    
    def _parse_value(self, value_str: str) -> Any:
        """値を解析（数値または文字列）"""
        # 数値に変換を試みる
        try:
            if '.' in value_str:
                return float(value_str)
            else:
                return int(value_str)
        except ValueError:
            # 数値でない場合は文字列として返す
            return value_str.strip('"\'')
    
    def _resolve_item(self, text: str) -> Optional[str]:
        """
        項目を解決
        
        優先順位:
        1. display_name完全一致
        2. key一致
        3. description中の略称一致
        """
        text = text.strip()
        
        # 1. display_name完全一致
        if text in self.item_key_by_display_name:
            return self.item_key_by_display_name[text]
        
        # 2. key一致
        if text in self.item_key_by_key:
            return self.item_key_by_key[text]
        
        # 3. description中の略称一致（部分一致）
        for item_key, item_def in self.item_dictionary.items():
            if isinstance(item_def, dict):
                description = item_def.get("description", "")
                if text in description:
                    return item_key
        
        return None
    
    def _resolve_event(self, text: str) -> Optional[str]:
        """イベントを解決"""
        text = text.strip()
        
        # name_jp一致
        if text in self.event_key_by_name_jp:
            return self.event_key_by_name_jp[text]
        
        # alias一致
        if text in self.event_key_by_alias:
            return self.event_key_by_alias[text]
        
        return None
    
    def _resolve_repro_metric(self, text: str) -> Optional[str]:
        """繁殖指標を解決"""
        text = text.strip()
        
        # 繁殖指標のマッピング
        repro_metrics = {
            "受胎率": "conception_rate",
            "発情発見率": "estrus_detection_rate",
            "妊娠率": "pregnancy_rate",
        }
        
        return repro_metrics.get(text, text)

