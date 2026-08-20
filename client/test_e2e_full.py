"""全流程联调：登录/项目/提交上传/密码/找回邮件/下载/只读查看/冲突。"""
import json
import time
import urllib.request
import urllib.error
import uuid
from pathlib import Path

CLIENT = "http://127.0.0.1:18792"
CLOUD = "http://127.0.0.1:18666"
SRV = Path("server")


def call(url, method="GET", data=None, headers=None, timeout=15):
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


def call_raw(url, method="GET", headers=None, timeout=15):
    req = urllib.request.Request(url, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as e:
        return e.code, e.read()


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
        with urllib.request.urlopen(req, timeout=15) as resp:
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


def login_c(pwd):
    return call(CLIENT + "/api/login", "POST", {"password": pwd})


# 1 管理员登录
s, r = login_c("admin123")
check("admin login", s == 200 and r.get("role") == "admin", str(r.get("role")))

# 2 项目列表（多 sheet 模板）
s, r = call(CLIENT + "/api/projects")
check("project list", s == 200 and len(r.get("projects", [])) >= 3, str(r.get("projects")))

# 3 首次同步（建立基线；云端此时可能无版本，先传一个）
s, r = call(CLIENT + "/api/sync-push", "POST", {})
check("initial push", s == 200 and r.get("baselineVersion"), str(r.get("baselineVersion")))

# 4 提交库存变动（同步上传）
projects = call(CLIENT + "/api/projects")[1].get("projects", [""])
s, r = call(CLIENT + "/api/submit", "POST", {
    "material": "联调测试电阻",
    "requester": "王五",
    "project": projects[0],
    "changeType": "入库",
    "quantity": "12",
    "specification": "E2E-001",
    "materialCode": "",
    "location": "货架A1",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "联调",
})
check("submit with sync upload", s == 200 and r.get("ok") and r.get("uploaded") is True,
      f"uploaded={r.get('uploaded')} err={r.get('uploadError')}")

# 5 新增项目（管理员）
s, r = call(CLIENT + "/api/projects", "POST", {"action": "add", "project": "联调项目"})
check("admin add project", s == 200 and "联调项目" in r.get("projects", []), str(r.get("projects")))

# 6 删除项目
s, r = call(CLIENT + "/api/projects", "POST", {"action": "remove", "project": "联调项目"})
check("admin remove project", s == 200 and "联调项目" not in r.get("projects", []), str(r.get("projects")))

# 7 添加普通用户并登录
s, r = call(CLIENT + "/api/users", "POST", {"action": "add", "password": "246810", "note": "联调用户"})
check("admin add user", s == 200 and r.get("ok"), str(r.get("ok")))
s, r = login_c("246810")
check("user login", s == 200 and r.get("role") == "user", str(r.get("role")))

# 8 普通用户提交（应成功并上传）
projects = call(CLIENT + "/api/projects")[1].get("projects", [""])
s, r = call(CLIENT + "/api/submit", "POST", {
    "material": "联调电容",
    "requester": "王五",
    "project": projects[0],
    "changeType": "入库",
    "quantity": "5",
    "specification": "E2E-002",
    "materialCode": "",
    "location": "货架A2",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "普通用户联调",
})
check("user submit + upload", s == 200 and r.get("ok") and r.get("uploaded") is True,
      f"uploaded={r.get('uploaded')}")

# 9 普通用户回溯被拒
s, r = call(CLIENT + "/api/rollback", "POST", {"versionId": "v-test"})
check("user rollback forbidden", s == 400 and "管理员" in str(r.get("error", "")), str(r.get("error")))

# 10 普通用户项目管理被拒
s, r = call(CLIENT + "/api/projects", "POST", {"action": "add", "project": "非法项目"})
check("user add project forbidden", s == 400 and "管理员" in str(r.get("error", "")), str(r.get("error")))

# 11 普通用户改密码被拒
s, r = call(CLIENT + "/api/admin/password", "POST", {"oldPassword": "admin123", "newPassword": "123456"})
check("user change password forbidden", s == 400 and "管理员" in str(r.get("error", "")), str(r.get("error")))

# 12 普通用户找回密码 → 提示联系管理员
s, r = call(CLIENT + "/api/forgot-password", "POST", {"type": "user"})
check("forgot user hint", s == 200 and "管理员" in str(r.get("message", "")), str(r.get("message")))

# 13 普通用户下载 Excel
s, raw = call_raw(CLIENT + "/api/download-excel")
check("user download excel", s == 200 and len(raw) > 10000, f"bytes={len(raw)}")

# 14 普通用户查看 Excel（只读副本）
s, r = call(CLIENT + "/api/view-excel", "POST", {})
check("user view excel copy", s == 200 and "只读" in str(r.get("message", "")), str(r.get("message")))
import openpyxl
copy_files = list(Path("view-copies").glob("库存查看副本-*.xlsx")) if Path("view-copies").is_dir() else []
protected = False
if copy_files:
    wb = openpyxl.load_workbook(copy_files[-1])
    protected = all(ws.protection.sheet for ws in wb.worksheets)
    wb.close()
check("view copy sheet protected", protected, f"files={len(copy_files)}")

# 15 管理员查看 Excel（直接打开正式文件，不产生副本行为校验跳过）

# 16 修改管理员密码 → 需重新登录（先切回管理员会话）
s, r = login_c("admin123")
check("relogin admin before pwd change", s == 200 and r.get("role") == "admin", str(r.get("role")))
s, r = call(CLIENT + "/api/admin/password", "POST", {"oldPassword": "admin123", "newPassword": "555888"})
check("admin change password", s == 200 and r.get("relogin"), str(r.get("ok")))
s, r = login_c("admin123")
check("old password dead", s == 400, str(r.get("error")))
s, r = login_c("555888")
check("new password login", s == 200 and r.get("role") == "admin", str(r.get("role")))

# 17 改回原密码
s, r = call(CLIENT + "/api/admin/password", "POST", {"oldPassword": "555888", "newPassword": "admin123"})
check("revert admin password", s == 200, str(r.get("ok")))
login_c("admin123")

# 18 密保邮箱
s, r = call(CLIENT + "/api/admin/email")
check("get security email", s == 200 and r.get("email") == "security@example.com", str(r.get("email")))
s, r = call(CLIENT + "/api/admin/email", "POST", {"email": "security@example.com"})
check("update security email", s == 200 and r.get("ok"), str(r.get("ok")))

# 19 找回管理员密码（假SMTP收邮件）
s, r = call(CLIENT + "/api/forgot-password", "POST", {"type": "admin"})
check("forgot admin sends email", s == 200 and "已发送" in str(r.get("message", "")), str(r.get("message")))
time.sleep(1)
inbox_files = list(SRV.glob("inbox/*.eml"))
check("email received in fake smtp", len(inbox_files) >= 1, f"files={len(inbox_files)}")

# 20 冲突场景：同事上传新版本 → 本机提交被拦
s, r = call(CLOUD + "/api/login", "POST", {"password": "admin123"}, {"Content-Type": "application/json"})
cloud_token = r.get("token", "")
s, r = cloud_upload(cloud_token, "assets/实验室库存管理模板.xlsx")
check("colleague uploads", s == 200 and r.get("ok"), str(r.get("versionId")))
s, r = call(CLIENT + "/api/submit", "POST", {
    "material": "冲突物料",
    "requester": "王五",
    "project": projects[0],
    "changeType": "入库",
    "quantity": "1",
    "specification": "CF-01",
    "materialCode": "",
    "location": "货架A1",
    "po": "",
    "changedAt": "2026-08-12",
    "note": "冲突",
})
check("submit blocked on conflict", s == 409 and r.get("conflict") is True, str(r.get("error")))

# 21 同步后恢复
s, r = call(CLIENT + "/api/sync-pull", "POST", {})
check("pull resolves conflict", s == 200 and r.get("baselineVersion"), str(r.get("baselineVersion")))

failed = [c for c in checks if not c[1]]
print(f"\n{len(checks) - len(failed)}/{len(checks)} passed")
raise SystemExit(1 if failed else 0)
