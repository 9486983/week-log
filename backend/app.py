# -*- coding: utf-8 -*-
"""WeekLog · Git 提交自动填充 SharePoint 周报。"""

import os
from datetime import datetime
from functools import wraps
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS

import excel_writer
import models
import services

FRONTEND_DIR = (Path(__file__).resolve().parent.parent / "frontend").resolve()

app = Flask(__name__, static_folder=None)
CORS(app)
app.json.ensure_ascii = False
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

models.init_db()


def api_route(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        try:
            return function(*args, **kwargs)
        except services.FillBusy as error:
            return jsonify({"error": str(error)}), 409
        except services.ValidationError as error:
            return jsonify({"error": str(error)}), 400
        except RuntimeError as error:
            return jsonify({"error": str(error)}), 500
        except Exception as error:  # noqa: BLE001
            app.logger.exception("API 处理失败")
            return jsonify({"error": "服务器内部错误: {}".format(error)}), 500

    return wrapper


@app.get("/health")
def health():
    return jsonify(
        {
            "status": "ok",
            "time": datetime.now().isoformat(sep=" ", timespec="seconds"),
            "excel_com": excel_writer._com_available(),
        }
    )


@app.get("/api/meta")
@api_route
def meta():
    return jsonify(
        {
            "name": "WeekLog",
            "version": "1.0.0",
            "excel_com": excel_writer._com_available(),
            "samples_dir": str(models.SAMPLES_DIR),
        }
    )


@app.get("/api/config")
@api_route
def get_config():
    return jsonify({"config": services.public_config()})


@app.put("/api/config")
@api_route
def put_config():
    payload = request.get_json(silent=True) or {}
    return jsonify({"config": services.update_config(payload)})


@app.get("/api/repos/status")
@api_route
def repos_status():
    import git_source

    repos = services.load_options()["repos"]
    return jsonify({"repos": [git_source.repo_status(path) for path in repos]})


@app.post("/api/repos/status")
@api_route
def repos_status_post():
    import git_source

    payload = request.get_json(silent=True) or {}
    repos = services.parse_repos(payload.get("repos")) or services.load_options()["repos"]
    return jsonify({"repos": [git_source.repo_status(path) for path in repos]})


@app.post("/api/collect")
@api_route
def collect():
    payload = request.get_json(silent=True) or {}
    preview = services.build_preview(
        payload.get("week_start"),
        author=payload.get("author"),
        repos=payload.get("repos"),
        start_day=payload.get("start_day"),
        range_start=payload.get("range_start"),
        range_end=payload.get("range_end"),
    )
    return jsonify(preview)


@app.post("/api/fill")
@api_route
def fill():
    payload = request.get_json(silent=True) or {}
    rows = payload.get("rows") or []
    dry_run = payload.get("dry_run")
    if dry_run is None:
        dry_run = services.load_options()["dry_run"]
    run_id = services.start_fill(
        payload.get("range_start") or payload.get("week_start"),
        rows,
        bool(dry_run),
        author=payload.get("author"),
        range_end=payload.get("range_end") or payload.get("week_end"),
    )
    return (
        jsonify(
            {
                "run_id": run_id,
                "mode": "live" if not dry_run else "dry_run",
                "message": "任务已启动",
            }
        ),
        202,
    )


@app.get("/api/runs")
@api_route
def runs():
    return jsonify(
        {
            "runs": services.list_runs(
                request.args.get("limit", 50),
                week_start=request.args.get("week_start"),
                week_end=request.args.get("week_end"),
                mode=request.args.get("mode"),
                status=request.args.get("status"),
            )
        }
    )


@app.get("/api/runs/<int:run_id>")
@api_route
def run_detail(run_id):
    data = services.run_detail(run_id)
    if data is None:
        return jsonify({"error": "填充记录不存在"}), 404
    return jsonify(data)


@app.post("/api/runs/<int:run_id>/retry")
@api_route
def retry(run_id):
    new_id = services.retry_run(run_id)
    return jsonify({"run_id": new_id, "message": "已重新启动"}), 202


@app.get("/api/runs/<int:run_id>/snapshot")
@api_route
def download_snapshot(run_id):
    data = services.run_detail(run_id)
    if not data or not data.get("snapshot_path"):
        return jsonify({"error": "该记录没有写入前快照"}), 404
    path = Path(data["snapshot_path"])
    if not path.exists():
        return jsonify({"error": "快照文件已不存在"}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.get("/api/runs/<int:run_id>/audit")
@api_route
def download_audit(run_id):
    data = services.run_detail(run_id)
    if not data or not data.get("audit_path"):
        return jsonify({"error": "该记录没有审计表"}), 404
    path = Path(data["audit_path"])
    if not path.exists():
        return jsonify({"error": "审计表已不存在"}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


@app.get("/api/samples")
@api_route
def samples():
    return jsonify({"samples": services.list_samples()})


@app.post("/api/samples/upload")
@api_route
def upload_sample():
    file = request.files.get("file")
    return jsonify({"sample": services.save_sample(file)})


@app.post("/api/template/inspect")
@api_route
def inspect_template():
    payload = request.get_json(silent=True) or {}
    path = payload.get("path") or services.load_options()["sample_path"]
    if not path:
        available = services.list_samples()
        if available:
            path = available[0]["path"]
    if not path:
        raise services.ValidationError("请先上传模板文件或填写本地模板路径")
    return jsonify(services.inspect_template(path))


@app.post("/api/sharepoint/inspect")
@api_route
def inspect_sharepoint():
    return jsonify(services.inspect_sharepoint())


@app.get("/api/workbook/status")
@api_route
def workbook_status():
    return jsonify(services.workbook_status())


@app.post("/api/ai/test")
@api_route
def ai_test():
    return jsonify(services.test_ai())


@app.post("/api/ai/rewrite")
@api_route
def ai_rewrite_rows():
    payload = request.get_json(silent=True) or {}
    return jsonify(services.rewrite_ai(payload.get("rows") or []))


@app.get("/")
def serve_index():
    return send_from_directory(FRONTEND_DIR, "index.html")


@app.get("/<path:path>")
def serve_spa(path):
    if path.startswith("api/"):
        return jsonify({"error": "Not Found"}), 404
    target = (FRONTEND_DIR / path).resolve()
    if FRONTEND_DIR in target.parents and target.is_file():
        return send_from_directory(FRONTEND_DIR, path)
    if "." in Path(path).name:
        return jsonify({"error": "Not Found"}), 404
    return send_from_directory(FRONTEND_DIR, "index.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    print("WeekLog 已启动: http://localhost:{}".format(port))
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
