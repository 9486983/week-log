# -*- coding: utf-8 -*-
"""数据库模型与配置读写。"""

import json
import os
from datetime import datetime
from pathlib import Path

from sqlalchemy import (
    Boolean,
    Column,
    Date,
    DateTime,
    ForeignKey,
    Integer,
    String,
    Text,
    create_engine,
    event,
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
SNAPSHOT_DIR = DATA_DIR / "snapshots"
AUDIT_DIR = DATA_DIR / "audit"
LOG_DIR = DATA_DIR / "logs"
SAMPLES_DIR = BASE_DIR.parent / "samples"

for _directory in (DATA_DIR, SNAPSHOT_DIR, AUDIT_DIR, LOG_DIR, SAMPLES_DIR):
    _directory.mkdir(parents=True, exist_ok=True)

DATABASE_URL = os.environ.get("DATABASE_URL") or "sqlite:///{}".format(
    (DATA_DIR / "app.db").as_posix()
)
_CONNECT_ARGS = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=_CONNECT_ARGS, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)
Base = declarative_base()


if DATABASE_URL.startswith("sqlite"):

    @event.listens_for(engine, "connect")
    def _sqlite_pragmas(dbapi_connection, _connection_record):
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.close()


def _loads_list(raw):
    if not raw:
        return []
    try:
        value = json.loads(raw)
        return value if isinstance(value, list) else [str(value)]
    except (ValueError, TypeError):
        return [str(raw)]


def _iso(value):
    return value.isoformat(sep=" ", timespec="seconds") if value else None


class FillRun(Base):
    __tablename__ = "fill_runs"

    id = Column(Integer, primary_key=True)
    created_at = Column(DateTime, default=datetime.now, nullable=False)
    started_at = Column(DateTime)
    finished_at = Column(DateTime)
    week_start = Column(Date)
    week_end = Column(Date)
    author = Column(String(120))
    repos = Column(Text)
    mode = Column(String(20), default="dry_run")
    status = Column(String(20), default="pending")
    message = Column(Text)
    target_url = Column(Text)
    sheet_name = Column(String(200))
    header_row = Column(Integer)
    start_row = Column(Integer)
    end_row = Column(Integer)
    rows_count = Column(Integer, default=0)
    commits_count = Column(Integer, default=0)
    snapshot_path = Column(Text)
    audit_path = Column(Text)
    log_path = Column(Text)

    rows = relationship(
        "FillRow",
        back_populates="run",
        cascade="all, delete-orphan",
        order_by="FillRow.seq",
    )
    commits = relationship(
        "FillCommit",
        back_populates="run",
        cascade="all, delete-orphan",
        order_by="FillCommit.committed_at",
    )

    def to_dict(self, include_children=False):
        data = {
            "id": self.id,
            "created_at": _iso(self.created_at),
            "started_at": _iso(self.started_at),
            "finished_at": _iso(self.finished_at),
            "week_start": self.week_start.isoformat() if self.week_start else None,
            "week_end": self.week_end.isoformat() if self.week_end else None,
            "range_start": self.week_start.isoformat() if self.week_start else None,
            "range_end": self.week_end.isoformat() if self.week_end else None,
            "author": self.author,
            "repos": _loads_list(self.repos),
            "mode": self.mode,
            "status": self.status,
            "message": self.message,
            "target_url": self.target_url,
            "sheet_name": self.sheet_name,
            "header_row": self.header_row,
            "start_row": self.start_row,
            "end_row": self.end_row,
            "rows_count": self.rows_count or 0,
            "commits_count": self.commits_count or 0,
            "snapshot_path": self.snapshot_path,
            "audit_path": self.audit_path,
            "log_path": self.log_path,
            "has_snapshot": bool(self.snapshot_path and Path(self.snapshot_path).exists()),
            "has_audit": bool(self.audit_path and Path(self.audit_path).exists()),
        }
        if include_children:
            data["rows"] = [row.to_dict() for row in self.rows]
            data["commits"] = [commit.to_dict() for commit in self.commits]
        return data


class FillRow(Base):
    __tablename__ = "fill_rows"

    id = Column(Integer, primary_key=True)
    run_id = Column(Integer, ForeignKey("fill_runs.id", ondelete="CASCADE"), index=True)
    seq = Column(Integer, default=0)
    ticket = Column(String(120))
    description = Column(Text)
    repos = Column(Text)
    branches = Column(Text)
    date_range = Column(String(80))
    commit_count = Column(Integer, default=0)
    hours = Column(String(40))
    progress = Column(String(40))
    remark = Column(Text)
    excel_row = Column(Integer)
    written = Column(Boolean, default=False)
    error = Column(Text)

    run = relationship("FillRun", back_populates="rows")

    def to_dict(self):
        return {
            "id": self.id,
            "seq": self.seq,
            "ticket": self.ticket,
            "description": self.description,
            "repos": self.repos,
            "branches": self.branches,
            "date_range": self.date_range,
            "commit_count": self.commit_count or 0,
            "hours": self.hours,
            "progress": self.progress,
            "remark": self.remark,
            "excel_row": self.excel_row,
            "written": bool(self.written),
            "error": self.error,
        }


class FillCommit(Base):
    __tablename__ = "fill_commits"

    id = Column(Integer, primary_key=True)
    run_id = Column(Integer, ForeignKey("fill_runs.id", ondelete="CASCADE"), index=True)
    commit_hash = Column(String(64))
    author = Column(String(120))
    committed_at = Column(DateTime)
    repo = Column(String(200))
    branch = Column(String(200))
    subject = Column(Text)
    ticket = Column(String(120))

    run = relationship("FillRun", back_populates="commits")

    def to_dict(self):
        return {
            "commit_hash": self.commit_hash,
            "author": self.author,
            "committed_at": _iso(self.committed_at),
            "repo": self.repo,
            "branch": self.branch,
            "subject": self.subject,
            "ticket": self.ticket,
        }


class Config(Base):
    __tablename__ = "config"

    id = Column(Integer, primary_key=True)
    config_key = Column(String(80), unique=True, nullable=False)
    config_value = Column(Text)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)


SHAREPOINT_URL = ""

DEFAULT_REPOS = []

DEFAULT_ALIASES = {
    "ticket": ["需求/Ticket编号", "Ticket/需求编号", "需求编号", "Ticket编号", "Ticket", "需求号", "需求"],
    "description": ["任务描述", "本周完成事项", "完成事项", "工作内容", "任务内容", "工作描述", "描述", "任务"],
    "author": ["执行人", "姓名", "人员", "填报人"],
    "repo": ["代码路径", "仓库", "项目", "代码库"],
    "branch": ["分支", "分支名"],
    "date": ["日期", "提交日期", "时间范围"],
    "commits": ["提交数", "提交次数", "提交记录数"],
    "hours": ["投入工时(h)", "投入工时", "工时(h)", "工时", "工作量(h)"],
    "progress": ["进度(%)", "进度", "完成度"],
    "remark": ["备注", "状态", "风险/问题", "说明"],
}

DEFAULT_AI_PROMPT = """你是周报助手。根据同一个 Ticket 下的 Git 提交记录，写一条简洁、专业的中文任务描述，用于周报的「任务描述」列。
要求：
1. 合并同类项，不要逐条罗列 commit；
2. 保留关键业务信息（需求号、模块/页面、修复或开发内容）；
3. 1-3 句话，不超过 120 字；
4. 不要编造提交记录里没有的内容；
5. 只输出描述正文，不要加引号、编号或前缀。"""

DEFAULT_CONFIGS = {
    "sharepoint_url": SHAREPOINT_URL,
    "write_mode": "open",
    "target_sheet": "",
    "target_mode": "existing",
    "weekly_template_sheet": "",
    "weekly_sheet_prefix": "周报汇总",
    "weekly_overwrite": "1",
    "header_row": "0",
    "author": "",
    "author_display": "",
    "repos": json.dumps(DEFAULT_REPOS, ensure_ascii=False),
    "week_start_day": "0",
    "footer_keywords": "合计,总计,小计",
    "column_aliases": json.dumps(DEFAULT_ALIASES, ensure_ascii=False),
    "date_mode": "last",
    "aggregate_mode": "daily",
    "reuse_empty_rows": "1",
    "dry_run": "1",
    "copy_row_format": "1",
    "insert_before_footer": "1",
    "ignore_merge_commits": "1",
    "max_description_chars": "1000",
    "sample_path": "",
    "ai_enabled": "0",
    "ai_base_url": "https://opencode.ai/zen/v1",
    "ai_api_key": "",
    "ai_model": "deepseek-v4.1-flash",
    "ai_prompt": DEFAULT_AI_PROMPT,
    "ai_timeout": "60",
    "ai_temperature": "0.3",
}


def init_db():
    Base.metadata.create_all(engine)
    with SessionLocal() as session:
        existing = {row.config_key for row in session.query(Config).all()}
        for key, value in DEFAULT_CONFIGS.items():
            if key not in existing:
                session.add(Config(config_key=key, config_value=value))
        stale = (
            session.query(FillRun)
            .filter(FillRun.status.in_(("pending", "running")))
            .all()
        )
        for run in stale:
            run.status = "failed"
            run.message = "服务重启，任务中断（可重试）"
            run.finished_at = datetime.now()
        session.commit()


def get_config(key, default=None):
    with SessionLocal() as session:
        row = session.query(Config).filter(Config.config_key == key).first()
        if row is None or row.config_value is None:
            return DEFAULT_CONFIGS.get(key, default)
        return row.config_value


def set_config(key, value):
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
    with SessionLocal() as session:
        row = session.query(Config).filter(Config.config_key == key).first()
        if row is None:
            session.add(Config(config_key=key, config_value=text))
        else:
            row.config_value = text
        session.commit()


def get_all_configs():
    result = dict(DEFAULT_CONFIGS)
    with SessionLocal() as session:
        for row in session.query(Config).all():
            if row.config_value is not None:
                result[row.config_key] = row.config_value
    return result
