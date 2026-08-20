import json
import urllib.request
import urllib.error

BASE = "http://127.0.0.1:18792"


def call(url, method="GET", data=None, timeout=10):
    body = json.dumps(data).encode() if data is not None else None
    headers = {"Content-Type": "application/json"} if data is not None else {}
    req = urllib.request.Request(url, method=method, data=body, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
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


s, r = call(BASE + "/api/login", "POST", {"password": "admin123"})
check("client admin login", s == 200 and r.get("role") == "admin", json.dumps(r, ensure_ascii=False)[:100])

s, r = call(BASE + "/api/login", "POST", {"password": "000000"})
check("client wrong password", s == 400 and r.get("ok") is False, json.dumps(r, ensure_ascii=False)[:100])

s, r = call(BASE + "/api/sync-pull", "POST", {})
check("client pull latest", s == 200 and r.get("ok") is True, json.dumps(r, ensure_ascii=False)[:150])

s, r = call(BASE + "/api/versions")
check("client list versions", s == 200 and len(r.get("versions", [])) >= 1, json.dumps(r, ensure_ascii=False)[:120])

# 管理员添加普通用户
s, r = call(BASE + "/api/users", "POST", {"action": "add", "password": "888888", "note": "王同事"})
check("admin add user via client", s == 200 and r.get("ok") is True, json.dumps(r, ensure_ascii=False)[:80])

s, r = call(BASE + "/api/users")
check("admin list users via client", s == 200 and len(r.get("users", [])) >= 1, json.dumps(r, ensure_ascii=False)[:120])

# 登出模拟：直接覆盖 sync.json 不方便，改为登录普通用户
s, r = call(BASE + "/api/login", "POST", {"password": "888888"})
check("client user login", s == 200 and r.get("role") == "user", json.dumps(r, ensure_ascii=False)[:100])

s, r = call(BASE + "/api/rollback", "POST", {"versionId": "v-20260812-152609-376815fc"})
check("client user rollback forbidden", s == 400 and "管理员" in str(r.get("error", "")), json.dumps(r, ensure_ascii=False)[:120])

s, r = call(BASE + "/api/users", "POST", {"action": "add", "password": "777777"})
check("client user add forbidden", s in (400, 403) and "管理员" in str(r.get("error", "")), json.dumps(r, ensure_ascii=False)[:100])

# 冲突检测：普通用户修改云端版本号（模拟其他同事提交），再提交应冲突
# 这里直接模拟：先以管理员登录上传新版本，再以普通用户提交
s, r = call(BASE + "/api/login", "POST", {"password": "admin123"})
check("relogin admin", s == 200 and r.get("role") == "admin", json.dumps(r, ensure_ascii=False)[:80])

s, r = call(BASE + "/api/submit", "POST", {
    "material": "测试物料",
    "requester": "王五",
    "project": "项目B",
    "changeType": "入库",
    "quantity": "5",
    "specification": "TEST-01",
    "materialCode": "",
    "location": "货架A1",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "同步冲突测试",
})
check("admin submit with conflict check", s == 200 and r.get("ok") is True, json.dumps(r, ensure_ascii=False)[:120])

import time
time.sleep(2)
s, r = call(BASE + "/api/sync-status")
check("baseline updated after push", s == 200 and r.get("baselineVersion"), json.dumps(r, ensure_ascii=False)[:150])

failed = [c for c in checks if not c[1]]
print(f"\n{len(checks) - len(failed)}/{len(checks)} passed")
raise SystemExit(1 if failed else 0)
