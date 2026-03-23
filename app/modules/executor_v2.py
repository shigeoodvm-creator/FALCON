"""
FALCON2 - Executor V2（改修版）
ExecutionPlanに従い、LIST/AGG/EVENTCOUNT/GRAPH/REPROを実行
"""

import logging
from typing import Dict, Any, Optional, List
from pathlib import Path
from datetime import datetime, timedelta

from db.db_handler import DBHandler
from modules.formula_engine import FormulaEngine
from modules.rule_engine import RuleEngine
from modules.execution_plan import ExecutionPlan, AnalysisType, GraphType, Condition, GroupBy
from modules.query_router_v2 import QueryRouterV2
from app.settings_manager import SettingsManager

logger = logging.getLogger(__name__)


class ExecutorV2:
    """
    ExecutionPlanに従って分析を実行
    
    実行種別:
    - LIST: リスト表示
    - AGG: 集計
    - EVENTCOUNT: イベント集計
    - GRAPH: グラフ
    - REPRO: 繁殖分析
    """
    
    def __init__(self, db_handler: DBHandler, formula_engine: FormulaEngine,
                 rule_engine: RuleEngine, item_dictionary_path: Optional[Path] = None,
                 event_dictionary_path: Optional[Path] = None, farm_path: Optional[Path] = None):
        """
        初期化
        
        Args:
            db_handler: DBHandler インスタンス
            formula_engine: FormulaEngine インスタンス
            rule_engine: RuleEngine インスタンス
            item_dictionary_path: item_dictionary.json のパス
            event_dictionary_path: event_dictionary.json のパス
            farm_path: 農場フォルダパス（VWP取得用）
        """
        self.db = db_handler
        self.formula_engine = formula_engine
        self.rule_engine = rule_engine
        self.item_dictionary_path = item_dictionary_path
        self.event_dictionary_path = event_dictionary_path
        self.farm_path = farm_path
        
        # QueryRouterV2のイベントグルーピングを参照
        self.event_groups = QueryRouterV2.EVENT_GROUPS
        
        # SettingsManager（VWP取得用）
        self.settings_manager = None
        self.DEFAULT_VWP = 50  # デフォルトVWP（日数）
        if farm_path:
            try:
                self.settings_manager = SettingsManager(farm_path)
                self.settings_manager.load()
            except Exception as e:
                logger.warning(f"SettingsManager初期化エラー: {e}")
    
    def execute(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """
        ExecutionPlanを実行
        
        Args:
            plan: ExecutionPlan
        
        Returns:
            実行結果の辞書
            {
                "success": bool,
                "data": Any,  # データ（表、グラフデータ等）
                "errors": List[str],
                "warnings": List[str]
            }
        """
        # エラーチェック
        if plan.errors:
            return {
                "success": False,
                "data": None,
                "errors": plan.errors,
                "warnings": plan.warnings
            }
        
        # 分析種別ごとに実行
        if plan.analysis_type == AnalysisType.LIST:
            return self._execute_list(plan)
        elif plan.analysis_type == AnalysisType.AGG:
            return self._execute_agg(plan)
        elif plan.analysis_type == AnalysisType.EVENTCOUNT:
            return self._execute_eventcount(plan)
        elif plan.analysis_type == AnalysisType.GRAPH:
            return self._execute_graph(plan)
        elif plan.analysis_type == AnalysisType.REPRO:
            return self._execute_repro(plan)
        else:
            return {
                "success": False,
                "data": None,
                "errors": [f"未対応の分析種別: {plan.analysis_type}"],
                "warnings": []
            }
    
    def _execute_list(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """リスト実行"""
        try:
            # 全牛を取得
            cows = self.db.get_all_cows()
            
            # 条件でフィルタリング
            filtered_cows = self._filter_cows(cows, plan.conditions)
            
            # 期間でフィルタリング（必要に応じて）
            if plan.period_start and plan.period_end:
                # 期間フィルタリングは後で実装（必要に応じて）
                pass
            
            # 各牛の項目を計算
            results = []
            for cow in filtered_cows:
                cow_auto_id = cow.get("auto_id")
                if not cow_auto_id:
                    continue
                
                row = {}
                for column_key in plan.columns:
                    # 項目を計算
                    calculated = self.formula_engine.calculate(cow_auto_id, item_key=column_key)
                    value = calculated.get(column_key, "")
                    row[column_key] = value
                
                if row:
                    row["cow_id"] = cow.get("cow_id", "")
                    row["auto_id"] = cow_auto_id
                    results.append(row)
            
            return {
                "success": True,
                "data": {
                    "type": "table",
                    "columns": plan.columns,
                    "rows": results
                },
                "errors": [],
                "warnings": plan.warnings
            }
        
        except Exception as e:
            logger.error(f"リスト実行エラー: {e}", exc_info=True)
            return {
                "success": False,
                "data": None,
                "errors": [f"リスト実行エラー: {str(e)}"],
                "warnings": plan.warnings
            }
    
    def _execute_agg(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """集計実行"""
        try:
            # 全牛を取得
            cows = self.db.get_all_cows()
            
            # 条件でフィルタリング
            filtered_cows = self._filter_cows(cows, plan.conditions)
            
            # GROUP BY処理
            if plan.group_by:
                # GROUP BYありの集計
                grouped_results = self._group_and_aggregate(filtered_cows, plan.group_by, plan.aggregate_items)
                return {
                    "success": True,
                    "data": {
                        "type": "table",
                        "columns": ["group"] + plan.aggregate_items,
                        "rows": grouped_results
                    },
                    "errors": [],
                    "warnings": plan.warnings
                }
            else:
                # GROUP BYなしの集計（全体集計）
                aggregated_results = self._aggregate(filtered_cows, plan.aggregate_items)
                return {
                    "success": True,
                    "data": {
                        "type": "table",
                        "columns": plan.aggregate_items,
                        "rows": [aggregated_results]
                    },
                    "errors": [],
                    "warnings": plan.warnings
                }
        
        except Exception as e:
            logger.error(f"集計実行エラー: {e}", exc_info=True)
            return {
                "success": False,
                "data": None,
                "errors": [f"集計実行エラー: {str(e)}"],
                "warnings": plan.warnings
            }
    
    def _execute_eventcount(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """イベント集計実行"""
        try:
            # イベントキーからイベント番号を取得
            event_numbers = []
            for event_key in plan.event_keys:
                if event_key in self.event_groups:
                    # グルーピング
                    event_numbers.extend(self.event_groups[event_key])
                else:
                    # 個別イベント（event_dictionaryから取得する必要がある）
                    # 簡易実装：event_keyを数値として解釈
                    try:
                        event_numbers.append(int(event_key))
                    except ValueError:
                        logger.warning(f"イベントキーを解決できませんでした: {event_key}")
            
            if not event_numbers:
                return {
                    "success": False,
                    "data": None,
                    "errors": ["イベント番号を解決できませんでした"],
                    "warnings": plan.warnings
                }
            
            # イベントを取得
            events = []
            for event_number in event_numbers:
                if plan.period_start and plan.period_end:
                    period_events = self.db.get_events_by_number_and_period(
                        event_number, plan.period_start, plan.period_end
                    )
                else:
                    period_events = self.db.get_events_by_number(event_number)
                events.extend(period_events)
            
            # GROUP BY処理
            if plan.group_by:
                grouped_results = self._group_event_count(events, plan.group_by)
                return {
                    "success": True,
                    "data": {
                        "type": "table",
                        "columns": ["group", "count"],
                        "rows": grouped_results
                    },
                    "errors": [],
                    "warnings": plan.warnings
                }
            else:
                # GROUP BYなし（全体集計）
                count = len(events)
                return {
                    "success": True,
                    "data": {
                        "type": "table",
                        "columns": ["count"],
                        "rows": [{"count": count}]
                    },
                    "errors": [],
                    "warnings": plan.warnings
                }
        
        except Exception as e:
            logger.error(f"イベント集計実行エラー: {e}", exc_info=True)
            return {
                "success": False,
                "data": None,
                "errors": [f"イベント集計実行エラー: {str(e)}"],
                "warnings": plan.warnings
            }
    
    def _execute_graph(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """グラフ実行"""
        # 簡易実装：後で実装
        return {
            "success": False,
            "data": None,
            "errors": ["グラフ機能は未実装です"],
            "warnings": plan.warnings
        }
    
    def _execute_repro(self, plan: ExecutionPlan) -> Dict[str, Any]:
        """繁殖分析実行（DC305互換）"""
        try:
            # VWPを取得（農場設定から、デフォルト50）
            vwp = self.DEFAULT_VWP
            if self.settings_manager:
                try:
                    settings = self.settings_manager.load()
                    # 農場設定からVWPを取得（将来的に実装予定の「農場繁殖指標設定」から取得）
                    # 現時点ではデフォルト値を使用
                    vwp = settings.get('reproduction', {}).get('vwp', self.DEFAULT_VWP)
                except Exception as e:
                    logger.warning(f"VWP取得エラー: {e}、デフォルト値({self.DEFAULT_VWP})を使用")
            
            # 繁殖分析を実行
            repro_analysis = ReproductionAnalysis(
                db=self.db,
                rule_engine=self.rule_engine,
                formula_engine=self.formula_engine,
                vwp=vwp
            )
            
            results = repro_analysis.analyze(
                period_start=plan.period_start,
                period_end=plan.period_end
            )
            
            # DC305互換のテーブル形式に変換
            rows = []
            total_br_el = 0
            total_bred = 0
            total_preg = 0
            
            for result in results:
                rows.append({
                    "cycle": f"Cycle {result.cycle_number}",
                    "start_date": result.start_date,
                    "end_date": result.end_date,
                    "br_el": result.br_el,
                    "bred": result.bred,
                    "preg": result.preg,
                    "hdr": result.hdr,
                    "pr": result.pr
                })
                total_br_el += result.br_el
                total_bred += result.bred
                total_preg += result.preg
            
            # Total/Average行を追加
            avg_hdr = (total_bred / total_br_el * 100) if total_br_el > 0 else 0.0
            avg_pr = (total_preg / total_br_el * 100) if total_br_el > 0 else 0.0
            
            rows.append({
                "cycle": "Total/Average",
                "start_date": "",
                "end_date": "",
                "br_el": total_br_el,
                "bred": total_bred,
                "preg": total_preg,
                "hdr": round(avg_hdr, 2),
                "pr": round(avg_pr, 2)
            })
            
            return {
                "success": True,
                "data": {
                    "type": "table",
                    "columns": ["cycle", "start_date", "end_date", "br_el", "bred", "preg", "hdr", "pr"],
                    "rows": rows
                },
                "errors": [],
                "warnings": plan.warnings
            }
        
        except Exception as e:
            logger.error(f"繁殖分析実行エラー: {e}", exc_info=True)
            return {
                "success": False,
                "data": None,
                "errors": [f"繁殖分析実行エラー: {str(e)}"],
                "warnings": plan.warnings
            }
    
    def _filter_cows(self, cows: List[Dict[str, Any]], conditions: List[Condition]) -> List[Dict[str, Any]]:
        """条件で牛をフィルタリング"""
        filtered = []
        
        for cow in cows:
            cow_auto_id = cow.get("auto_id")
            if not cow_auto_id:
                continue
            
            # すべての条件を満たすかチェック
            matches_all = True
            for condition in conditions:
                if not self._check_condition(cow, cow_auto_id, condition):
                    matches_all = False
                    break
            
            if matches_all:
                filtered.append(cow)
        
        return filtered
    
    def _check_condition(self, cow: Dict[str, Any], cow_auto_id: int, condition: Condition) -> bool:
        """条件をチェック"""
        item_key = condition.item_key
        
        # 項目の値を取得
        if item_key in cow:
            # cowテーブルに直接存在する項目
            value = cow.get(item_key)
        else:
            # 計算項目
            calculated = self.formula_engine.calculate(cow_auto_id, item_key=item_key)
            value = calculated.get(item_key)
        
        if value is None:
            return False
        
        # 演算子で比較
        operator = condition.operator
        target_value = condition.value
        
        if operator == "=":
            return value == target_value
        elif operator == "!=":
            return value != target_value
        elif operator == ">":
            return value > target_value
        elif operator == ">=":
            return value >= target_value
        elif operator == "<":
            return value < target_value
        elif operator == "<=":
            return value <= target_value
        else:
            return False
    
    def _group_and_aggregate(self, cows: List[Dict[str, Any]], group_by: GroupBy,
                            aggregate_items: List[str]) -> List[Dict[str, Any]]:
        """GROUP BYして集計"""
        # グループごとに集計
        groups: Dict[str, List[Dict[str, Any]]] = {}
        
        for cow in cows:
            cow_auto_id = cow.get("auto_id")
            if not cow_auto_id:
                continue
            
            # グループキーを取得
            if group_by.bin_base and group_by.bin_width:
                # BIN処理（DIM区分など）
                calculated = self.formula_engine.calculate(cow_auto_id, item_key=group_by.bin_base)
                value = calculated.get(group_by.bin_base)
                if value is not None:
                    bin_index = int(value) // group_by.bin_width
                    group_key = f"{bin_index * group_by.bin_width}-{(bin_index + 1) * group_by.bin_width - 1}"
                else:
                    group_key = "不明"
            elif group_by.item_key:
                # 通常の項目でグループ化
                if group_by.item_key in cow:
                    group_key = str(cow.get(group_by.item_key))
                else:
                    calculated = self.formula_engine.calculate(cow_auto_id, item_key=group_by.item_key)
                    group_key = str(calculated.get(group_by.item_key, "不明"))
            else:
                group_key = "全体"
            
            if group_key not in groups:
                groups[group_key] = []
            groups[group_key].append(cow)
        
        # 各グループで集計
        results = []
        for group_key, group_cows in sorted(groups.items()):
            aggregated = self._aggregate(group_cows, aggregate_items)
            aggregated["group"] = group_key
            results.append(aggregated)
        
        return results
    
    def _aggregate(self, cows: List[Dict[str, Any]], aggregate_items: List[str]) -> Dict[str, Any]:
        """全体集計"""
        result = {}
        
        for item_key in aggregate_items:
            values = []
            for cow in cows:
                cow_auto_id = cow.get("auto_id")
                if not cow_auto_id:
                    continue
                
                # 項目を計算
                calculated = self.formula_engine.calculate(cow_auto_id, item_key=item_key)
                value = calculated.get(item_key)
                
                if value is not None:
                    try:
                        # 数値に変換可能な場合のみ集計対象
                        if isinstance(value, (int, float)):
                            values.append(float(value))
                    except (ValueError, TypeError):
                        pass
            
            if values:
                # 平均値を計算（簡易実装：後で他の集計関数も追加可能）
                result[item_key] = sum(values) / len(values)
            else:
                result[item_key] = None
        
        return result
    
    def _group_event_count(self, events: List[Dict[str, Any]], group_by: GroupBy) -> List[Dict[str, Any]]:
        """イベントをGROUP BYして集計"""
        # グループごとに集計
        groups: Dict[str, int] = {}
        
        for event in events:
            cow_auto_id = event.get("cow_auto_id")
            if not cow_auto_id:
                continue
            
            # グループキーを取得
            if group_by.bin_base and group_by.bin_width:
                # BIN処理（DIM区分など）
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                if cow:
                    calculated = self.formula_engine.calculate(cow_auto_id, item_key=group_by.bin_base)
                    value = calculated.get(group_by.bin_base)
                    if value is not None:
                        bin_index = int(value) // group_by.bin_width
                        group_key = f"{bin_index * group_by.bin_width}-{(bin_index + 1) * group_by.bin_width - 1}"
                    else:
                        group_key = "不明"
                else:
                    group_key = "不明"
            elif group_by.item_key:
                # 通常の項目でグループ化
                cow = self.db.get_cow_by_auto_id(cow_auto_id)
                if cow:
                    if group_by.item_key in cow:
                        group_key = str(cow.get(group_by.item_key))
                    else:
                        calculated = self.formula_engine.calculate(cow_auto_id, item_key=group_by.item_key)
                        group_key = str(calculated.get(group_by.item_key, "不明"))
                else:
                    group_key = "不明"
            else:
                group_key = "全体"
            
            groups[group_key] = groups.get(group_key, 0) + 1
        
        # 結果をリストに変換
        results = []
        for group_key, count in sorted(groups.items()):
            results.append({"group": group_key, "count": count})
        
        return results

