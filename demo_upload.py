#!/usr/bin/env python3
"""
DingTalk File Upload Demo Script
Tests the full upload chain: Get Token → Get Space → Upload File → Get fileId
Usage: python demo_upload.py /path/to/file.xlsx
"""

import sys
import os
import json
import hashlib
import requests
from pathlib import Path
from dotenv import load_dotenv

ENV_PATH = Path(__file__).parent / ".env"
if not ENV_PATH.exists():
    print("❌ .env not found in project dir, copying from PaySignPrinter...")
    os.system("cp /home/ubuntu/coding/PaySignPrinter/.env " + str(ENV_PATH))
    print("✅ .env copied.")

load_dotenv(ENV_PATH)

APP_KEY = os.getenv("DINGTALK_APP_KEY")
APP_SECRET = os.getenv("DINGTALK_APP_SECRET")
AGENT_ID = os.getenv("DINGTALK_AGENT_ID")
USER_ID = os.getenv("DINGTALK_USER_ID")
UNION_ID = os.getenv("DINGTALK_UNION_ID")

if not all([APP_KEY, APP_SECRET, AGENT_ID]):
    print("❌ Missing credentials in .env file")
    sys.exit(1)

def print_step(name):
    print(f"\n{'='*60}")
    print(f"▶ {name}")
    print('='*60)

def check_response(resp, step_name):
    if resp.status_code >= 400:
        print(f"❌ {step_name} failed: HTTP {resp.status_code}")
        try:
            print(json.dumps(resp.json(), indent=2, ensure_ascii=False))
        except Exception:
            print(resp.text)
        return False
    print(f"✅ {step_name} succeeded (HTTP {resp.status_code})")
    return True

def main():
    user_id = USER_ID
    union_id = UNION_ID
    file_arg = None

    i = 1
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg in ("--user-id", "-u") and i + 1 < len(sys.argv):
            i += 1
            user_id = sys.argv[i]
        elif arg in ("--union-id",) and i + 1 < len(sys.argv):
            i += 1
            union_id = sys.argv[i]
        elif not arg.startswith("-") and file_arg is None:
            file_arg = arg
        i += 1

    if not file_arg:
        print("Usage: python demo_upload.py <path_to_excel_file> [--user-id USER_ID] [--union-id UNION_ID]")
        print("\nNote: userId is required for Step 2 (Get Attachment Space).")
        print("      unionId is required for Step 3 (Get Upload Info).")
        print("Provide them via:")
        print("  1. --user-id / --union-id CLI flags")
        print("  2. DINGTALK_USER_ID / DINGTALK_UNION_ID environment variables in .env")
        sys.exit(1)

    file_path = Path(file_arg)
    if not file_path.exists():
        print(f"❌ File not found: {file_path}")
        sys.exit(1)

    file_name = file_path.name
    file_size = file_path.stat().st_size
    file_type = "xlsx"

    with open(file_path, "rb") as f:
        file_data = f.read()
    file_md5 = hashlib.md5(file_data).hexdigest()

    print(f"📄 File: {file_name}")
    print(f"📦 Size: {file_size} bytes")
    print(f"🔐 MD5:  {file_md5}")
    if user_id:
        print(f"👤 userId: {user_id}")
    if union_id:
        print(f"👤 unionId: {union_id}")
    elif user_id:
        union_id = user_id
        print(f"👤 unionId: {union_id} (inherited from userId)")

    print_step("Step 1: Get AccessToken")
    token_url = "https://api.dingtalk.com/v1.0/oauth2/accessToken"
    token_payload = {"appKey": APP_KEY, "appSecret": APP_SECRET}
    resp = requests.post(token_url, json=token_payload, timeout=30)
    if not check_response(resp, "Get AccessToken"):
        sys.exit(1)
    token_data = resp.json()
    access_token = token_data.get("accessToken")
    print(f"   accessToken: {access_token[:20]}... (expires in {token_data.get('expireIn')}s)")

    headers = {"x-acs-dingtalk-access-token": access_token}

    print_step("Step 2: Get Attachment Space (获取审批钉盘空间信息)")
    if not user_id:
        print("❌ userId is required for this API but was not provided.")
        print("   Set DINGTALK_USER_ID in .env or pass --user-id <id>")
        sys.exit(1)

    space_url = "https://api.dingtalk.com/v1.0/workflow/processInstances/spaces/infos/query"
    space_payload = {"agentId": AGENT_ID, "userId": user_id}
    resp = requests.post(space_url, headers=headers, json=space_payload, timeout=30)
    if not check_response(resp, "Get Space"):
        sys.exit(1)
    space_data = resp.json()
    space_id = space_data.get("result", {}).get("spaceId")
    if not space_id:
        print("❌ spaceId not found in response")
        print(json.dumps(space_data, indent=2, ensure_ascii=False))
        sys.exit(1)
    print(f"   spaceId: {space_id}")

    print_step("Step 3: Get Upload Info (获取文件上传信息)")
    if not union_id:
        print("❌ unionId is required for this API but was not provided.")
        print("   Set DINGTALK_UNION_ID in .env or pass --union-id <id>")
        sys.exit(1)

    upload_info_url = (
        f"https://api.dingtalk.com/v1.0/drive/spaces/{space_id}/files/0/uploadInfos"
        f"?unionId={union_id}&fileName={requests.utils.quote(file_name)}"
        f"&fileSize={file_size}&md5={file_md5}"
    )
    resp = requests.get(upload_info_url, headers=headers, timeout=30)
    if not check_response(resp, "Get Upload Info"):
        sys.exit(1)
    upload_info = resp.json()
    print(json.dumps(upload_info, indent=2, ensure_ascii=False))

    header_info = upload_info.get("headerSignatureUploadInfo", {})
    resource_url = header_info.get("resourceUrl")
    upload_headers = header_info.get("headers", {})
    media_id = header_info.get("mediaId")
    upload_key = upload_info.get("uploadKey")

    if not resource_url:
        print("❌ resourceUrl not found in upload info")
        sys.exit(1)

    print_step("Step 4: Upload File to OSS")
    print(f"   PUT {resource_url[:80]}...")
    print(f"   Headers: {upload_headers}")

    resp = requests.put(resource_url, headers=upload_headers, data=file_data, timeout=120)
    if not check_response(resp, "OSS Upload"):
        sys.exit(1)
    print(f"   Response body: {resp.text}")

    file_id = None
    if upload_key:
        print_step("Step 5: Commit File (if uploadKey exists)")
        commit_url = "https://api.dingtalk.com/v2.0/storage/spaces/files/0/commit"
        commit_payload = {"uploadKey": upload_key, "spaceId": space_id}
        resp = requests.post(commit_url, headers=headers, json=commit_payload, timeout=30)
        if not check_response(resp, "Commit File"):
            print("   Warning: Commit failed, but upload may still be usable.")
        else:
            commit_data = resp.json()
            print(json.dumps(commit_data, indent=2, ensure_ascii=False))
            file_id = commit_data.get("id") or commit_data.get("fileId") or commit_data.get("unionId")
    else:
        print("\n⚠️  No uploadKey returned, skipping commit step.")

    print_step("Summary & DDAttachment Format Reference")
    print("""
DDAttachment value format (JSON string):
{
  "spaceId": "{space_id}",
  "fileId": "{file_id}",
  "fileName": "{file_name}",
  "fileSize": {file_size},
  "fileType": "{file_type}"
}
""")
    print(f"   spaceId : {space_id}")
    print(f"   fileId  : {file_id or media_id or '(unknown — check commit response above)'}")
    print(f"   fileName: {file_name}")
    print(f"   fileSize: {file_size}")
    print(f"   fileType: {file_type}")
    print(f"   mediaId : {media_id or 'N/A'}")
    print(f"   uploadKey: {upload_key or 'N/A'}")

    if file_id:
        attachment_json = json.dumps({
            "spaceId": space_id,
            "fileId": file_id,
            "fileName": file_name,
            "fileSize": file_size,
            "fileType": file_type
        }, ensure_ascii=False)
        print(f"\n📎 DDAttachment JSON:\n{attachment_json}")
    else:
        print("\n⚠️  fileId not resolved. Use mediaId or check API docs for exact field.")

    print("\n✅ Demo completed.")


if __name__ == "__main__":
    main()
