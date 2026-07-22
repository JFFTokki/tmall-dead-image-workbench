from __future__ import annotations

import argparse
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path


BASE_URL = "https://tb.maijsoft.cn/index.php"
DEFAULT_FIELDS = ["num_iid", "pic1", "pic2", "rectangle_pic1", "rectangle_pic2"]


@dataclass
class SessionInfo:
    cookie: str
    user_agent: str
    referer: str


def read_ids(path: Path | None, inline_ids: list[str]) -> list[str]:
    raw: list[str] = []
    if path:
        raw.extend(path.read_text(encoding="utf-8-sig").splitlines())
    raw.extend(inline_ids)

    ids: list[str] = []
    seen: set[str] = set()
    for line in raw:
        for value in re.split(r"[\s,，]+", line.strip()):
            if not value:
                continue
            if value not in seen:
                seen.add(value)
                ids.append(value)
    return ids


def chunked(values: list[str], size: int) -> list[list[str]]:
    return [values[index : index + size] for index in range(0, len(values), size)]


def session_from_har(path: Path, cookie_override: str = "") -> SessionInfo:
    har = json.loads(path.read_text(encoding="utf-8"))
    entries = har.get("log", {}).get("entries", [])

    chosen = None
    for entry in entries:
        request = entry.get("request", {})
        post_params = request.get("postData", {}).get("params", [])
        route = next((p.get("value") for p in post_params if p.get("name") == "r"), "")
        if route in {"export%2FtaskStart", "export/taskStart"}:
            chosen = request
            break
    if chosen is None:
        raise ValueError("HAR 中没有找到 export/taskStart 请求，请重新抓取一次开始导出的请求。")

    headers = {h["name"].lower(): h.get("value", "") for h in chosen.get("headers", [])}
    cookie = cookie_override.strip() or headers.get("cookie", "")
    if not cookie:
        raise ValueError("没有 Cookie，无法复用登录状态。请使用 --cookie 或 --cookie-file。")

    return SessionInfo(
        cookie=cookie,
        user_agent=headers.get("user-agent", "Mozilla/5.0"),
        referer=headers.get(
            "referer",
            "https://tb.maijsoft.cn/index.php?r=export%2Findex",
        ),
    )


def request_json(
    url: str,
    session: SessionInfo,
    *,
    data: list[tuple[str, str]] | None = None,
    timeout: int = 60,
) -> dict:
    body = None
    headers = {
        "User-Agent": session.user_agent,
        "Cookie": session.cookie,
        "Referer": session.referer,
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    }
    if data is not None:
        body = urllib.parse.urlencode(data).encode("utf-8")
        headers["Content-Type"] = "application/x-www-form-urlencoded; charset=UTF-8"

    request = urllib.request.Request(url, data=body, headers=headers, method="POST" if data else "GET")
    with urllib.request.urlopen(request, timeout=timeout) as response:
        text = response.read().decode("utf-8", errors="replace")
    return json.loads(text)


def build_export_params(ids: list[str], route: str, fields: list[str]) -> list[tuple[str, str]]:
    params = [
        ("r", route),
        ("itemStatus[]", "onsale"),
        ("instockStatus", "for_shelved"),
        ("itemCids", ""),
        ("keyword", ""),
        ("startPrice", ""),
        ("endPrice", ""),
        ("startModified", ""),
        ("endModified", ""),
        ("itemIds", ",".join(ids)),
    ]
    params.extend(("field[]", field) for field in fields)
    params.extend([("imgSize", "120"), ("skuImgSize", "60")])
    return params


def wait_for_download_url(session: SessionInfo, start_result: dict, poll_interval: float = 2.0) -> str:
    progress_id = start_result["progressId"]
    user_id = start_result["userId"]

    for attempt in range(120):
        last = "true" if attempt > 0 else "false"
        query = urllib.parse.urlencode(
            {
                "r": "export/getProgress",
                "last": last,
                "progressId": progress_id,
                "userId": user_id,
                "editType": "",
                "logDir": "",
                "errorRow": "3" if last == "true" else "1",
                "skipRow": "1",
                "msgRow": "1",
                "_": str(int(time.time() * 1000)),
            }
        )
        result = request_json(f"{BASE_URL}?{query}", session)
        for item in result.get("errors", []):
            if isinstance(item, dict) and item.get("type") == "downURL":
                return item["value"]
        if str(result.get("status")) == "2":
            # Finished but the download URL may appear only in the last=true poll.
            time.sleep(poll_interval)
        else:
            time.sleep(poll_interval)

    raise TimeoutError(f"导出任务超时，progressId={progress_id}")


def download_file(url: str, session: SessionInfo, output_path: Path) -> None:
    headers = {
        "User-Agent": session.user_agent,
        "Cookie": session.cookie,
        "Referer": session.referer,
        "Accept": "*/*",
    }
    request = urllib.request.Request(url, headers=headers)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(request, timeout=180) as response:
        output_path.write_bytes(response.read())


def export_batch(session: SessionInfo, ids: list[str], fields: list[str], output_dir: Path, batch_no: int) -> Path:
    check_params = build_export_params(ids, "export/taskCheck", fields)
    check_result = request_json(BASE_URL, session, data=check_params)
    if not check_result.get("success"):
        raise RuntimeError(f"导出检查失败：{check_result}")

    start_params = build_export_params(ids, "export/taskStart", fields)
    start_result = request_json(BASE_URL, session, data=start_params)
    if not start_result.get("success"):
        raise RuntimeError(f"导出启动失败：{start_result}")

    download_url = wait_for_download_url(session, start_result)
    name = download_url.rsplit("/", 1)[-1] or f"maijsoft_export_batch_{batch_no}.xlsx"
    output_path = output_dir / f"{batch_no:03d}_{name}"
    download_file(download_url, session, output_path)
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(description="通过掌柜软件接口按商品ID批量导出图片元数据 Excel")
    parser.add_argument("--har", type=Path, default=Path("tb.maijsoft.cn.har"), help="从已登录浏览器导出的 HAR 文件")
    parser.add_argument("--cookie", default="", help="已登录浏览器中的 Cookie 字符串")
    parser.add_argument("--cookie-file", type=Path, help="保存 Cookie 字符串的文本文件")
    parser.add_argument("--ids-file", type=Path, help="商品ID文本文件，每行一个或用逗号/空格分隔")
    parser.add_argument("--id", action="append", default=[], help="单个或多个商品ID；可重复传入")
    parser.add_argument("--output-dir", type=Path, default=Path("maijsoft_exports"), help="Excel 下载目录")
    parser.add_argument("--batch-size", type=int, default=10000, help="每批商品ID数量，最大 10000")
    parser.add_argument("--include-extra-fields", action="store_true", help="附加导出标题/价格/库存/商家编码")
    parser.add_argument("--dry-run", action="store_true", help="只显示分批和字段，不提交导出")
    args = parser.parse_args()

    if args.batch_size < 1 or args.batch_size > 10000:
        raise ValueError("--batch-size 必须在 1 到 10000 之间")

    ids = read_ids(args.ids_file, args.id)
    if not ids:
        raise ValueError("没有读取到商品ID，请使用 --ids-file 或 --id")

    fields = list(DEFAULT_FIELDS)
    if args.include_extra_fields:
        fields.extend(["title", "price", "num", "outer_id"])

    batches = chunked(ids, args.batch_size)
    print(f"商品ID总数：{len(ids)}，分批：{len(batches)}，每批最多：{args.batch_size}")
    print(f"导出字段：{','.join(fields)}；图片尺寸：120x120")

    if args.dry_run:
        for index, batch in enumerate(batches, start=1):
            print(f"批次 {index}: {len(batch)} 个，首个ID {batch[0]}，末个ID {batch[-1]}")
        return 0

    cookie = args.cookie
    if args.cookie_file:
        cookie = args.cookie_file.read_text(encoding="utf-8").strip()

    session = session_from_har(args.har, cookie_override=cookie)
    for index, batch in enumerate(batches, start=1):
        print(f"开始导出批次 {index}/{len(batches)}：{len(batch)} 个ID")
        try:
            output_path = export_batch(session, batch, fields, args.output_dir, index)
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f"HTTP {exc.code}：可能登录已过期，请重新导出 HAR。") from exc
        print(f"已下载：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
