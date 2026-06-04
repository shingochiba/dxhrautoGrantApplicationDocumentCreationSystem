"""
前回入力した会社情報・社労士情報をローカルJSONに永続化するモジュール。
ログインユーザーごとに別ファイルで保存する。
"""
from __future__ import annotations

import json
from dataclasses import asdict, fields
from pathlib import Path
from typing import List, Optional

from .company_form import CompanyInfo, SocialInsuranceLabor, Office, TrainingCompany
from . import auth


# プロジェクトルート直下の data フォルダに保存
_BASE_DIR = Path(__file__).resolve().parent.parent
_STORAGE_DIR = _BASE_DIR / "data"
# config フォルダ（マスタデータ等の Git 管理ファイル）
_CONFIG_DIR = _BASE_DIR / "config"
_SR_MASTER_FILE = _CONFIG_DIR / "sr_master.json"
_INDUSTRY_CODES_FILE = _CONFIG_DIR / "industry_codes.json"
_TRAINING_COMPANY_MASTER_FILE = _CONFIG_DIR / "training_company_master.json"


def _get_storage_file() -> Path:
    """ログインユーザーごとの保存ファイルパスを返す"""
    user = auth.get_current_user()
    if user:
        user_id = auth.get_user_id(user)
        return _STORAGE_DIR / f"last_inputs_{user_id}.json"
    # 未ログイン時（ローカル開発等）のフォールバック
    return _STORAGE_DIR / "last_inputs.json"


def _load_raw() -> dict:
    path = _get_storage_file()
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_raw(data: dict) -> None:
    _STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    with _get_storage_file().open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _filter_fields(data: dict, cls) -> dict:
    """dataclassに定義されていないキーを除去（後方互換）"""
    valid = {f.name for f in fields(cls)}
    return {k: v for k, v in data.items() if k in valid}


def load_saved_company() -> Optional[CompanyInfo]:
    raw = _load_raw().get("company")
    if not raw:
        return None
    try:
        # offices はネストしたリストなので分離して復元
        raw_offices = raw.get("offices") or []
        offices = []
        for o in raw_offices:
            if isinstance(o, dict):
                offices.append(Office(**_filter_fields(o, Office)))
        filtered = _filter_fields(raw, CompanyInfo)
        filtered["offices"] = offices
        return CompanyInfo(**filtered)
    except TypeError:
        return None


def save_company(company: CompanyInfo) -> None:
    data = _load_raw()
    data["company"] = asdict(company)
    _save_raw(data)


def load_saved_sr() -> Optional[SocialInsuranceLabor]:
    raw = _load_raw().get("sr")
    if not raw:
        return None
    try:
        return SocialInsuranceLabor(**_filter_fields(raw, SocialInsuranceLabor))
    except TypeError:
        return None


def save_sr(sr: SocialInsuranceLabor) -> None:
    data = _load_raw()
    data["sr"] = asdict(sr)
    _save_raw(data)


def clear_saved() -> None:
    path = _get_storage_file()
    if path.exists():
        path.unlink()


def load_industry_codes() -> dict:
    """日本標準産業分類（中分類2桁）コード→分類名の辞書を読み込む。
    ファイルが無い・壊れている場合は空辞書を返す。
    """
    if not _INDUSTRY_CODES_FILE.exists():
        return {}
    try:
        with _INDUSTRY_CODES_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {}
        # メタデータのキー（_comment 等）は除外し、2桁コードのみ返す
        return {k: v for k, v in data.items()
                if isinstance(k, str) and k.isdigit() and len(k) == 2}
    except (json.JSONDecodeError, OSError):
        return {}


def load_sr_master() -> List[SocialInsuranceLabor]:
    """社労士マスタを config/sr_master.json から読み込む。
    ファイルが無い・壊れている場合は空リストを返す。
    """
    if not _SR_MASTER_FILE.exists():
        return []
    try:
        with _SR_MASTER_FILE.open("r", encoding="utf-8") as f:
            raw_list = json.load(f)
        if not isinstance(raw_list, list):
            return []
        result: List[SocialInsuranceLabor] = []
        for item in raw_list:
            if isinstance(item, dict):
                result.append(SocialInsuranceLabor(**_filter_fields(item, SocialInsuranceLabor)))
        return result
    except (json.JSONDecodeError, OSError, TypeError):
        return []


def load_training_company_master() -> List[TrainingCompany]:
    """教育訓練機関 (研修実施会社) のマスタを config/training_company_master.json から読み込む。
    ファイルが無い・壊れている場合は空リストを返す。
    """
    if not _TRAINING_COMPANY_MASTER_FILE.exists():
        return []
    try:
        with _TRAINING_COMPANY_MASTER_FILE.open("r", encoding="utf-8") as f:
            raw_list = json.load(f)
        if not isinstance(raw_list, list):
            return []
        result: List[TrainingCompany] = []
        for item in raw_list:
            if isinstance(item, dict):
                result.append(TrainingCompany(**_filter_fields(item, TrainingCompany)))
        return result
    except (json.JSONDecodeError, OSError, TypeError):
        return []
