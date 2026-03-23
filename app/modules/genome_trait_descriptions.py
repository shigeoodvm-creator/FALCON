"""
FALCON2 - ゲノム項目の用語解説（レポート・選択UIでユーザーが参照できるようにする）
NotebookGM の traitDictionary および乳牛ゲノム評価の一般的な説明に基づく。
"""

from typing import Dict, Any, List, Tuple

# 部門の表示順（項目選択・用語解説のグループ表示に使用）
DEPARTMENT_ORDER: List[str] = [
    "経済価値",
    "乳用",
    "飼料効率",
    "繁殖",
    "耐久性・生存",
    "体型",
    "分娩",
    "その他",
]

# FALCONの項目キー -> { name_ja, desc, tip, direction, department }
# direction: higher=高いほど良い, lower=低いほど良い, neutral=色分けしない
# department: 部門分類（項目選択を部門ごとに表示するために使用）
GENOME_TRAIT_DESCRIPTIONS: Dict[str, Dict[str, Any]] = {
    "GDPR": {
        "name_ja": "GDPR（Daughter Pregnancy Rate）",
        "desc": "娘牛妊娠率のゲノム評価値。繁殖指標。",
        "tip": "高いほど繁殖成績が良いと評価されます。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GNM$": {
        "name_ja": "GNM$（Net Merit）",
        "desc": "純利益指標。乳用牛の総合的な経済的価値をドル建てで表したゲノム評価値。",
        "tip": "高いほど収益性が高いと評価されます。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GTPI": {
        "name_ja": "GTPI（Total Performance Index）",
        "desc": "総合能力指数。産乳・体型・耐久性・繁殖などを総合したゲノム評価値。",
        "tip": "高いほど総合的に優れています。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GDWP$": {
        "name_ja": "GDWP$",
        "desc": "DWP$ のゲノム評価値。乳量・乳成分・耐久性等の経済価値指標。",
        "tip": "高いほど経済的に有利です。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GCM$": {
        "name_ja": "GCM$",
        "desc": "Cheese Merit。チーズ用乳の経済価値。",
        "tip": "高いほどチーズ生産に有利です。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GFM$": {
        "name_ja": "GFM$",
        "desc": "Fluid Merit。飲用乳の経済価値。",
        "tip": "高いほど飲用乳生産に有利です。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GGM$": {
        "name_ja": "GGM$",
        "desc": "Grazing Merit。放牧適性の経済価値。",
        "tip": "高いほど放牧経営に適しています。",
        "direction": "higher",
        "department": "経済価値",
    },
    "GCA$": {
        "name_ja": "GCA$",
        "desc": "Calving Ability。分娩能力指数。分娩容易性と死産率を組み合わせた経済価値。",
        "tip": "高いほど難産・死産が少ないと評価されます。",
        "direction": "higher",
        "department": "経済価値",
    },
    "G乳量": {
        "name_ja": "G乳量",
        "desc": "乳量の遺伝的能力（ゲノム評価値）。",
        "tip": "高いほど乳量の遺伝的潜在能力が高いと評価されます。",
        "direction": "higher",
        "department": "乳用",
    },
    "GFAT": {
        "name_ja": "GFAT",
        "desc": "乳脂率・乳脂量の遺伝的能力。",
        "tip": "高いほど乳脂の遺伝的潜在能力が高いと評価されます。",
        "direction": "higher",
        "department": "乳用",
    },
    "GPROT": {
        "name_ja": "GPROT",
        "desc": "乳蛋白率・乳蛋白量の遺伝的能力。",
        "tip": "高いほど乳蛋白の遺伝的潜在能力が高いと評価されます。",
        "direction": "higher",
        "department": "乳用",
    },
    "GSCS": {
        "name_ja": "GSCS",
        "desc": "体細胞スコア。低いほど乳房炎リスクが低い。",
        "tip": "低いほど良い指標です。",
        "direction": "lower",
        "department": "乳用",
    },
    "GRFI": {
        "name_ja": "GRFI",
        "desc": "残差飼料摂取量。低いほど飼料効率が良い。",
        "tip": "低いほど良い指標です。",
        "direction": "lower",
        "department": "飼料効率",
    },
    "GFE": {
        "name_ja": "GFE",
        "desc": "飼料効率の遺伝的能力。",
        "tip": "高いほど飼料効率が良いと評価されます。",
        "direction": "higher",
        "department": "飼料効率",
    },
    "GHCR": {
        "name_ja": "GHCR",
        "desc": "処女牛（育成牛）の受胎率。高いほど受胎しやすい。",
        "tip": "Heifer Conception Rate。高いほど繁殖成績が良いと評価されます。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GCCR": {
        "name_ja": "GCCR",
        "desc": "成牛の受胎率。泌乳牛の受胎のしやすさの遺伝的能力。",
        "tip": "Cow Conception Rate。高いほど受胎しやすいと評価されます。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GFI": {
        "name_ja": "GFI",
        "desc": "繁殖指数または分娩容易性。データソースにより意味が異なります。",
        "tip": "高いほど良いとされることが多いです。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GEFC": {
        "name_ja": "GEFC",
        "desc": "Early First Calving。初産年齢の若さ。若い初産を遺伝的に持つほどプラス。",
        "tip": "プラスほど娘牛の初産が早くなり、育成コスト削減に寄与します。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GMSPD": {
        "name_ja": "GMSPD",
        "desc": "多胎率の遺伝的能力。",
        "tip": "データソースにより解釈が異なります。",
        "direction": "higher",
        "department": "繁殖",
    },
    "GPL": {
        "name_ja": "GPL",
        "desc": "Productive Life。持久力・生産寿命。",
        "tip": "高いほど長く生産に寄与する遺伝的能力が高いと評価されます。",
        "direction": "higher",
        "department": "耐久性・生存",
    },
    "GLIV": {
        "name_ja": "GLIV",
        "desc": "生存率の遺伝的能力。",
        "tip": "高いほど生存率が高いと評価されます。",
        "direction": "higher",
        "department": "耐久性・生存",
    },
    "GFS": {
        "name_ja": "GFS（体型総合）",
        "desc": "体型スコアの総合。乳房・肢蹄・体深などの総合評価。",
        "tip": "高いほど体型が優れていると評価されます。",
        "direction": "higher",
        "department": "体型",
    },
    "GTYPE FS": {
        "name_ja": "GTYPE FS",
        "desc": "体型総合（Type Final Score）のゲノム評価値。",
        "tip": "高いほど体型が優れていると評価されます。",
        "direction": "higher",
        "department": "体型",
    },
    "GUDC": {
        "name_ja": "GUDC",
        "desc": "乳房複合。乳房の深さ・付着の遺伝的能力。",
        "tip": "高いほど乳房形態が優れていると評価されます。",
        "direction": "higher",
        "department": "体型",
    },
    "GBDC": {
        "name_ja": "GBDC",
        "desc": "体深複合。体深の遺伝的能力。",
        "tip": "標準では色分けしません。",
        "direction": "neutral",
        "department": "体型",
    },
    "GFLC": {
        "name_ja": "GFLC",
        "desc": "肢蹄複合。肢蹄の遺伝的能力。",
        "tip": "高いほど肢蹄が優れていると評価されます。",
        "direction": "higher",
        "department": "体型",
    },
    "GSCE": {
        "name_ja": "GSCE",
        "desc": "種雄牛の分娩容易性。低いほど難産が少ない。",
        "tip": "Sire Calving Ease。低いほど良い指標です。",
        "direction": "lower",
        "department": "分娩",
    },
    "GDCE": {
        "name_ja": "GDCE",
        "desc": "娘牛の分娩容易性。低いほど難産が少ない。",
        "tip": "Daughter Calving Ease。低いほど良い指標です。",
        "direction": "lower",
        "department": "分娩",
    },
    "GSSB": {
        "name_ja": "GSSB",
        "desc": "単胎時の死産率。低いほど良い。",
        "tip": "Stillbirth (Single)。低いほど良い指標です。",
        "direction": "lower",
        "department": "分娩",
    },
    "GDSB": {
        "name_ja": "GDSB",
        "desc": "双胎時の死産率。低いほど良い。",
        "tip": "Stillbirth (Twin)。低いほど良い指標です。",
        "direction": "lower",
        "department": "分娩",
    },
    "GGL": {
        "name_ja": "GGL",
        "desc": "妊娠期間（日数）の遺伝的効果。",
        "tip": "標準では色分けしません。データソースにより意味が異なる場合があります。",
        "direction": "neutral",
        "department": "その他",
    },
    "GBVDV Results": {
        "name_ja": "GBVDV Results",
        "desc": "BVDV検査結果（牛ウイルス性下痢症）。",
        "tip": "色分け・グラフ対象外です。",
        "direction": "neutral",
        "department": "その他",
    },
    "GMSPD Reliability": {
        "name_ja": "GMSPD Reliability",
        "desc": "GMSPDの信頼性。",
        "tip": "標準では色分けしません。",
        "direction": "neutral",
        "department": "その他",
    },
}


def get_trait_description(key: str) -> Dict[str, Any]:
    """項目キーに対応する用語解説を返す。無い場合は display_name のみの簡易情報。"""
    return GENOME_TRAIT_DESCRIPTIONS.get(
        key,
        {"name_ja": key, "desc": "", "tip": "", "direction": "higher", "department": "その他"},
    )


def get_keys_grouped_by_department(keys: List[str]) -> List[Tuple[str, List[str]]]:
    """
    キー一覧を部門ごとにグループ化して返す。
    Returns:
        [(部門名, [key1, key2, ...]), ...] の形式。部門の並びは DEPARTMENT_ORDER に従う。
    """
    from collections import defaultdict
    dept_to_keys: Dict[str, List[str]] = defaultdict(list)
    for k in keys:
        dept = get_trait_description(k).get("department", "その他")
        dept_to_keys[dept].append(k)
    # 部門内は name_ja でソート
    for dept in dept_to_keys:
        dept_to_keys[dept].sort(key=lambda x: get_trait_description(x).get("name_ja", x))
    result: List[Tuple[str, List[str]]] = []
    for dept in DEPARTMENT_ORDER:
        if dept in dept_to_keys and dept_to_keys[dept]:
            result.append((dept, dept_to_keys[dept]))
    # 辞書にしかない部門（DEPARTMENT_ORDER に無い）は末尾に
    for dept in sorted(dept_to_keys.keys()):
        if dept not in DEPARTMENT_ORDER and dept_to_keys[dept]:
            result.append((dept, dept_to_keys[dept]))
    return result
