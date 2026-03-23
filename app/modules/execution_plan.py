"""
FALCON2 - ExecutionPlan（実行計画）
QueryRouterが生成する実行計画のデータ構造
"""

from typing import Dict, Any, Optional, List
from dataclasses import dataclass, field
from enum import Enum


class AnalysisType(Enum):
    """分析種別"""
    LIST = "list"           # リスト
    AGG = "agg"            # 集計
    EVENTCOUNT = "eventcount"  # イベント集計
    GRAPH = "graph"        # グラフ
    REPRO = "repro"        # 繁殖分析


class GraphType(Enum):
    """グラフ種類"""
    LINE = "line"          # 折れ線
    BAR = "bar"            # 棒
    SCATTER = "scatter"    # プロット
    BOX = "box"            # 箱ひげ
    SURVIVAL = "survival"  # 生存曲線


@dataclass
class Condition:
    """条件（WHERE句）"""
    item_key: str          # 項目キー（例: "DIM", "LACT"）
    operator: str          # 演算子（=, !=, >, >=, <, <=）
    value: Any             # 値（数値または文字列）
    keyword: Optional[str] = None  # キーワード条件（"妊娠牛", "空胎"等）


@dataclass
class GroupBy:
    """GROUP BY句"""
    item_key: Optional[str] = None  # 項目キー（例: "LACT", "CLVD"）
    bin_base: Optional[str] = None   # BINの基準項目（例: "DIM"）
    bin_width: Optional[int] = None  # BINの幅（例: 7, 14, 21, 30, 50）


@dataclass
class ExecutionPlan:
    """実行計画"""
    analysis_type: AnalysisType     # 分析種別
    period_start: Optional[str] = None  # 期間開始（YYYY-MM-DD）
    period_end: Optional[str] = None    # 期間終了（YYYY-MM-DD）
    
    # リスト用
    columns: List[str] = field(default_factory=list)  # 表示項目リスト（item_key）
    
    # 集計用
    aggregate_items: List[str] = field(default_factory=list)  # 集計項目（例: "授精回数", "初回授精日数"）
    group_by: Optional[GroupBy] = None  # 分けて見る基準
    
    # イベント集計用
    event_keys: List[str] = field(default_factory=list)  # 集計したいイベント（例: "授精", "分娩"）
    
    # グラフ用
    y_axis_item: Optional[str] = None  # 縦軸項目
    x_axis_item: Optional[str] = None  # 横軸項目
    graph_type: Optional[GraphType] = None  # グラフ種類
    split_by_item: Optional[str] = None  # 色・線を分ける基準
    
    # 繁殖分析用
    repro_metric: Optional[str] = None  # 繁殖指標（例: "受胎率", "発情発見率"）
    
    # 共通
    conditions: List[Condition] = field(default_factory=list)  # 絞り込み条件
    
    # エラー情報
    errors: List[str] = field(default_factory=list)  # エラーメッセージ
    warnings: List[str] = field(default_factory=list)  # 警告メッセージ

    @property
    def type(self) -> str:
        """UI互換: query_type として analysis_type の文字列値を返す"""
        return self.analysis_type.value

    @property
    def start(self) -> Optional[str]:
        """UI互換: 期間開始"""
        return self.period_start

    @property
    def end(self) -> Optional[str]:
        """UI互換: 期間終了"""
        return self.period_end

    @property
    def as_of_date(self) -> Optional[str]:
        """UI互換: 基準日（period_end を流用）"""
        return self.period_end

    def to_dict(self) -> Dict[str, Any]:
        """辞書形式に変換"""
        result = {
            "analysis_type": self.analysis_type.value,
            "period_start": self.period_start,
            "period_end": self.period_end,
            "columns": self.columns,
            "aggregate_items": self.aggregate_items,
            "event_keys": self.event_keys,
            "y_axis_item": self.y_axis_item,
            "x_axis_item": self.x_axis_item,
            "graph_type": self.graph_type.value if self.graph_type else None,
            "split_by_item": self.split_by_item,
            "repro_metric": self.repro_metric,
            "conditions": [
                {
                    "item_key": c.item_key,
                    "operator": c.operator,
                    "value": c.value,
                    "keyword": c.keyword
                }
                for c in self.conditions
            ],
            "errors": self.errors,
            "warnings": self.warnings
        }
        
        if self.group_by:
            result["group_by"] = {
                "item_key": self.group_by.item_key,
                "bin_base": self.group_by.bin_base,
                "bin_width": self.group_by.bin_width
            }
        
        return result

















