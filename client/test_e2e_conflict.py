import json
import time
import urllib.request
import urllib.error
import uuid

CLIENT = "http://127.0.0.1:18792"
CLOUD = "http://127.0.0.1:18666"


def call(url, method="GET", data=None, headers=None, timeout=10):
    body = json.dumps(data).encode() if data is not None else None
    hdrs = {"Content-Type": "application/json"} if data is not None else {}
    hdrs.update(headers or {})
    req = urllib.request.Request(url, method=method, data=body, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except ValueError:
            return e.code, {}


def cloud_upload(token, path):
    boundary = uuid.uuid4().hex
    file_data = open(path, "rb").read()
    parts = [
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"实验室库存管理.xlsx\"\r\nContent-Type: application/octet-stream\r\n\r\n".encode(),
        file_data,
        f"\r\n--{boundary}--\r\n".encode(),
    ]
    req = urllib.request.Request(
        CLOUD + "/api/upload",
        method="POST",
        data=b"".join(parts),
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}",
            "Authorization": f"Bearer {token}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode("utf-8"))
        except ValueError:
            return e.code, {}


checks = []


def check(name, ok, detail=""):
    checks.append((name, ok))
    print(("PASS " if ok else "FAIL ") + name + (" | " + detail if detail else ""))


# 1 管理员登录客户端
s, r = call(CLIENT + "/api/login", "POST", {"password": "admin123"})
check("login admin", s == 200 and r.get("role") == "admin", str(r.get("role")))

# 2 同步最新版本（建立基线）
s, r = call(CLIENT + "/api/sync-pull", "POST", {})
check("pull establishes baseline", s == 200 and r.get("baselineVersion"), str(r.get("baselineVersion")))

# 3 模拟同事A：云端直接上传新版本（内容不同）
s, r = call(CLOUD + "/api/login", "POST", {"password": "admin123"}, {"Content-Type": "application/json"})
cloud_token = r.get("token", "")
s, r = cloud_upload(cloud_token, "assets/实验室库存管理模板.xlsx")
check("colleague uploaded new version", s == 200 and r.get("ok"), str(r.get("versionId")))

# 4 本机提交 → 应被冲突拦截
s, r = call(CLIENT + "/api/submit", "POST", {
    "material": "冲突测试物料",
    "requester": "王五",
    "project": "项目B",
    "changeType": "入库",
    "quantity": "3",
    "specification": "CF-01",
    "materialCode": "",
    "location": "货架A1",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "e2e",
})
check("submit blocked on conflict", s == 409 and r.get("conflict") is True, json.dumps(r, ensure_ascii=False)[:100])

# 5 同步最新 → 基线更新
s, r = call(CLIENT + "/api/sync-pull", "POST", {})
check("pull resolves conflict", s == 200 and r.get("baselineVersion"), str(r.get("baselineVersion")))

# 6 再次提交 → 成功
s, r = call(CLIENT + "/api/submit", "POST", {
    "material": "冲突测试物料",
    "requester": "王五",
    "project": "项目B",
    "changeType": "入库",
    "quantity": "3",
    "specification": "CF-01",
    "materialCode": "",
    "location": "货架A1",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "e2e",
})
check("submit succeeds after pull", s == 200 and r.get("ok") is True, json.dumps(r, ensure_ascii=False)[:100])

# 7 等待异步上传完成，基线应前进
time.sleep(2)
s, r = call(CLIENT + "/api/sync-status")
check("auto push updated baseline", s == 200 and r.get("baselineVersion"), str(r.get("baselineVersion")))
check("no pending push", r.get("pendingPush") is False, f"pending={r.get('pendingPush')}")

failed = [c for c in checks if not c[1]]
print(f"\n{len(checks) - len(failed)}/{len(checks)} passed")
raise SystemExit(1 if failed else 0)
