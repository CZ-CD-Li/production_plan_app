"""应用级配置与路径定义。"""

import json
import os
import shutil
import sys
from pathlib import Path

APP_NAME = "生产计划协同管理平台"
APP_VERSION = "1.4.0"
ORG_NAME = "CraftWorks"

# 新版本安装程序发布目录（局域网共享），启动时自动检查
UPDATE_SHARE = r"\\192.168.8.14\工艺工程部\6-装配组\李翔\安装程序"
UPDATE_SHARE_ENV = "PRODUCTION_PLAN_UPDATE_SHARE"

DATA_FILES = ("app.db", "project_lead.db", "local_config.json")


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _writable(path: Path) -> bool:
    try:
        path.mkdir(parents=True, exist_ok=True)
        probe = path / ".write_test"
        probe.write_text("ok", encoding="utf-8")
        probe.unlink()
        return True
    except OSError:
        return False


def _local_app_data() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home())))


def _legacy_data_dirs() -> tuple[Path, ...]:
    """历史版本把数据写在 AppData 下，升级后需要迁回安装目录。"""
    base = _local_app_data()
    return (base / APP_NAME / "data", base / "ProductionPlan" / "data")


def _retarget_config(target: Path) -> None:
    """把仍指向旧 AppData 目录的数据库路径改到当前数据目录。"""
    config_file = target / "local_config.json"
    try:
        data = json.loads(config_file.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    if not isinstance(data, dict):
        return
    legacy_parents = {str(path).lower() for path in _legacy_data_dirs()}
    changed = False
    for key in ("db_path", "lead_db_path"):
        value = str(data.get(key, "")).strip()
        if not value:
            continue
        parent = Path(value).expanduser().parent
        if str(parent).lower() in legacy_parents:
            data[key] = str(target / Path(value).name)
            changed = True
    if changed:
        config_file.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )


def _migrate_legacy_data(target: Path) -> None:
    """首次运行时把旧 AppData 目录下的数据库复制到当前数据目录，避免数据“消失”。"""
    target.mkdir(parents=True, exist_ok=True)
    for legacy in _legacy_data_dirs():
        if not legacy.is_dir():
            continue
        try:
            if legacy.resolve() == target.resolve():
                continue
        except OSError:
            continue
        for name in DATA_FILES:
            source = legacy / name
            destination = target / name
            if not source.is_file() or destination.exists():
                continue
            try:
                shutil.copy2(source, destination)
            except OSError:
                continue
    _retarget_config(target)


if _is_frozen():
    # 打包后：数据目录固定在安装目录 data 下（安装程序默认装到当前用户目录，可写）；
    # 若安装目录不可写（如装在 Program Files），退回当前用户的 AppData。
    ROOT_DIR = Path(sys.executable).resolve().parent
    _portable_data = ROOT_DIR / "data"
    if _writable(_portable_data):
        DATA_DIR = _portable_data
    else:
        DATA_DIR = _local_app_data() / "ProductionPlan" / "data"
    _migrate_legacy_data(DATA_DIR)
else:
    ROOT_DIR = Path(__file__).resolve().parent.parent
    DATA_DIR = ROOT_DIR / "data"

APP_DIR = ROOT_DIR / "app"
DEMO_DIR = ROOT_DIR / "demo_data"
DB_PATH = DATA_DIR / "app.db"
LEAD_DB_PATH = DATA_DIR / "project_lead.db"
LOCAL_CONFIG_PATH = DATA_DIR / "local_config.json"
DB_PATH_ENV = "PRODUCTION_PLAN_DB"
LEAD_DB_PATH_ENV = "PRODUCTION_PLAN_LEAD_DB"

PROJECT_LEAD_SHEET = "新产品项目任务分布清单"
PLAN_FILE_KEYWORD = "301"

DATE_OFFSETS = (-2, -1, 0, 1, 2)
DATE_LABELS = {
    -2: "前天",
    -1: "昨天",
    0: "今天",
    1: "明天",
    2: "后天",
}
WEEKDAY_LABELS = ("周一", "周二", "周三", "周四", "周五", "周六", "周日")

MODES = ("装配", "焊接")
TABLE_HEADERS = ("生产线", "主机厂", "产品代号", "计划数量", "负责人", "备注")

# 测试阶段关闭 UDC 地址限制，允许保存到任意本地目录
ENFORCE_UDC_ADDRESS = False


def ensure_runtime_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not _is_frozen():
        DEMO_DIR.mkdir(parents=True, exist_ok=True)


def load_local_config() -> dict:
    """读取本机配置（数据库位置等），文件损坏时按默认值处理。"""
    try:
        data = json.loads(LOCAL_CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_local_config(values: dict) -> None:
    data = load_local_config()
    data.update(values)
    LOCAL_CONFIG_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def resolve_db_path() -> Path:
    """数据库位置优先级：环境变量 > 本机配置文件 > 默认本地路径。"""
    env_value = os.environ.get(DB_PATH_ENV, "").strip()
    if env_value:
        return Path(env_value).expanduser()
    configured = str(load_local_config().get("db_path", "")).strip()
    if configured:
        return Path(configured).expanduser()
    return DB_PATH


def resolve_lead_db_path() -> Path:
    """Project Lead 本机库位置：环境变量 > 本机配置文件 > data/project_lead.db。"""
    env_value = os.environ.get(LEAD_DB_PATH_ENV, "").strip()
    if env_value:
        return Path(env_value).expanduser()
    configured = str(load_local_config().get("lead_db_path", "")).strip()
    if configured:
        return Path(configured).expanduser()
    return LEAD_DB_PATH


def is_network_path(path: str | Path) -> bool:
    """判断是否为 UNC 路径或映射的网络驱动器。"""
    text = str(path)
    if text.startswith(("\\\\", "//")):
        return True
    drive, _ = os.path.splitdrive(text)
    if not drive or os.name != "nt":
        return False
    try:
        import ctypes

        return ctypes.windll.kernel32.GetDriveTypeW(f"{drive}\\") == 4
    except (AttributeError, OSError):
        return False
