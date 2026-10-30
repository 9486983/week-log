# -*- coding: utf-8 -*-
"""Git 提交采集与按 Ticket 汇总。

核心解析逻辑沿用 docs/git_commit_report.py，去掉了 Excel 导出部分，
改为返回结构化数据供 Web 端预览与写入。
"""

import re
import subprocess
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

DEFAULT_RANGE_DAYS = 30
GIT_TIMEOUT_SECONDS = 60

FIELD_SEPARATOR = "\x1f"
RECORD_SEPARATOR = "\x1e"

TICKET_RULES = [
    re.compile(r"\b([A-Z][A-Za-z0-9]*-\d{2,})\b"),
    re.compile(r"^\s*(\d{3,7})\b"),
    re.compile(r"[:：,，]\s*(\d{3,7})\b"),
    re.compile(r"^([A-Za-z][A-Za-z0-9 &/\-]{0,29}?)\s*#\d"),
    re.compile(r"^([A-Za-z][A-Za-z0-9 &/\-]{0,29}?):\s+\S"),
]

UNSORTED_TICKET = "未分类"


@dataclass(frozen=True)
class Commit:
    committed_at: datetime
    author: str
    subject: str
    body: str
    repo: str
    branch: str
    commit_hash: str

    def to_dict(self):
        return {
            "commit_hash": self.commit_hash,
            "author": self.author,
            "committed_at": self.committed_at.isoformat(sep=" ", timespec="seconds"),
            "repo": self.repo,
            "branch": self.branch,
            "subject": self.subject,
            "ticket": extract_ticket(self.subject) or extract_ticket(self.body),
        }


def run_git(repo_path, arguments):
    try:
        result = subprocess.run(
            ["git", *arguments],
            cwd=str(repo_path),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=GIT_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(
            "git 命令执行超时（>{}s）: git {}".format(
                GIT_TIMEOUT_SECONDS, " ".join(arguments)
            )
        )
    if result.returncode != 0:
        raise RuntimeError(
            "git 命令执行失败 [{}]: {}".format(repo_path, result.stderr.strip())
        )
    return result.stdout


def read_commit_log(repo_path):
    return run_git(
        repo_path,
        [
            "log",
            "--branches",
            "--remotes",
            "--tags",
            "--source",
            "--date=iso",
            "--pretty=format:%H{field}%an{field}%ad{field}%s{field}%b{field}%S{record}".format(
                field=FIELD_SEPARATOR,
                record=RECORD_SEPARATOR,
            ),
        ],
    )


def short_branch_name(source_ref):
    prefix = "refs/heads/"
    return source_ref[len(prefix):] if source_ref.startswith(prefix) else source_ref


def parse_commit_entry(entry, repo_name):
    parts = entry.split(FIELD_SEPARATOR)
    if len(parts) < 6:
        return None
    commit_hash, author, date_text, subject = parts[:4]
    body, source_ref = parts[4], parts[5]
    return Commit(
        committed_at=datetime.strptime(date_text[:19], "%Y-%m-%d %H:%M:%S"),
        author=author.strip(),
        subject=subject.strip(),
        body=body.strip(),
        repo=repo_name,
        branch=short_branch_name(source_ref.strip()),
        commit_hash=commit_hash.strip(),
    )


def extract_ticket(message):
    for rule in TICKET_RULES:
        match = rule.search(message)
        if match:
            return match.group(1)
    return ""


def load_commits(repo_paths, name, start_date, end_date, ignore_merges=False):
    parsed = {}
    for repo_path in repo_paths:
        repo_name = Path(repo_path).name
        raw_log = read_commit_log(repo_path)
        for entry in raw_log.split(RECORD_SEPARATOR):
            entry = entry.strip("\n")
            if not entry:
                continue
            commit = parse_commit_entry(entry, repo_name)
            if commit is not None:
                parsed[commit.commit_hash] = commit

    commits = parsed.values()
    if name:
        lowered = name.lower()
        commits = [c for c in commits if lowered in c.author.lower()]
    if ignore_merges:
        commits = [
            c for c in commits if not c.subject.lower().startswith("merge ")
        ]
    return sorted(
        (c for c in commits if start_date <= c.committed_at.date() <= end_date),
        key=lambda c: c.committed_at,
        reverse=True,
    )


def week_bounds(anchor, start_day=0):
    """返回 anchor 所在周的 (开始日, 结束日)。start_day: 0=周一, 6=周日。"""
    offset = (anchor.weekday() - int(start_day)) % 7
    start = anchor - timedelta(days=offset)
    return start, start + timedelta(days=6)


def _date_range_text(first_at, last_at):
    if first_at.date() == last_at.date():
        return first_at.strftime("%Y-%m-%d")
    return "{} ~ {}".format(first_at.strftime("%Y-%m-%d"), last_at.strftime("%Y-%m-%d"))


def aggregate_commits(commits, mode="daily", max_description_chars=1000):
    mode = (mode or "daily").strip().lower()
    buckets = {}
    for commit in commits:
        ticket = extract_ticket(commit.subject) or extract_ticket(commit.body)
        if mode == "ticket":
            key = (ticket,)
        else:
            key = (commit.committed_at.date().isoformat(), ticket)
        buckets.setdefault(key, []).append(commit)

    rows = []
    for key, items in buckets.items():
        items.sort(key=lambda c: c.committed_at)
        subjects = []
        for commit in items:
            subject = commit.subject.strip()
            if subject and subject not in subjects:
                subjects.append(subject)
        description = "；".join(subjects)
        if max_description_chars and len(description) > max_description_chars:
            description = description[: max_description_chars - 1] + "…"

        first_at = items[0].committed_at
        last_at = items[-1].committed_at
        if mode == "ticket":
            ticket = key[0]
            day = None
            date_range = _date_range_text(first_at, last_at)
            row_key = ticket or UNSORTED_TICKET
        else:
            day = key[0]
            ticket = key[1]
            date_range = day
            row_key = "{}|{}".format(day, ticket or UNSORTED_TICKET)
        rows.append(
            {
                "row_key": row_key,
                "ticket": ticket or UNSORTED_TICKET,
                "ticket_raw": ticket,
                "description": description,
                "subjects": subjects,
                "repos": "、".join(sorted({c.repo for c in items})),
                "branches": "、".join(sorted({c.branch for c in items})),
                "date_range": date_range,
                "day": day,
                "commit_count": len(items),
                "first_at": first_at.isoformat(sep=" ", timespec="seconds"),
                "last_at": last_at.isoformat(sep=" ", timespec="seconds"),
                "hours": "",
                "progress": "",
                "remark": "",
            }
        )

    if mode == "ticket":
        rows.sort(key=lambda row: (row["last_at"], row["ticket"]), reverse=True)
    else:
        rows.sort(key=lambda row: (row["day"] or "", row["ticket"]))
    return rows


def aggregate_by_ticket(commits, max_description_chars=1000):
    return aggregate_commits(commits, "ticket", max_description_chars)


def repo_status(path):
    repo = Path(path)
    info = {
        "path": str(repo),
        "name": repo.name,
        "exists": repo.exists(),
        "is_git": False,
        "last_commit": None,
        "error": None,
    }
    if not info["exists"]:
        info["error"] = "路径不存在"
        return info
    if not (repo / ".git").exists():
        info["error"] = "不是 Git 仓库（缺少 .git）"
        return info
    info["is_git"] = True
    try:
        output = run_git(
            repo,
            ["log", "-1", "--pretty=format:%h|%an|%ad|%s", "--date=iso"],
        ).strip()
        if output:
            short_hash, author, committed_at, subject = output.split("|", 3)
            info["last_commit"] = {
                "hash": short_hash,
                "author": author,
                "committed_at": committed_at[:19],
                "subject": subject,
            }
    except (RuntimeError, ValueError) as error:
        info["error"] = str(error)
    return info
