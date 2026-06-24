import requests
import os
import glob
import json
import sys
from datetime import datetime

# Force UTF-8 encoding for stdout on Windows to prevent UnicodeEncodeError
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Coca-Cola Zalo Mini App ID
COCA_APP_ID = "2284259347200678918"

def get_latest_coca_token():
    """Tự động tìm token Coca-Cola từ các file token được tạo bởi hook"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Tìm file có dạng token_MiniApp_2284259347200678918_*.json
    pattern = os.path.join(script_dir, f"token_MiniApp_{COCA_APP_ID}_*.json")
    files = glob.glob(pattern)
    if not files:
        return None, None
    
    # Lấy file mới nhất dựa trên thời gian sửa đổi
    latest_file = max(files, key=os.path.getmtime)
    try:
        with open(latest_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data, latest_file
    except Exception:
        return None, None

def main():
    print("=" * 60)
    print("      COCA-COLA TOKEN & PROFILE DETAILS GETTER")
    print("=" * 60)

    data, file_path = get_latest_coca_token()
    if not data:
        print(f"[ERR] Không tìm thấy file token cho Coca-Cola ({COCA_APP_ID})")
        print("Vui lòng kích hoạt Mini App Coca từ ứng dụng hoặc gửi broadcast để capture trước.")
        sys.exit(1)
        
    print(f"[OK] Đang đọc Token Coca từ file: {os.path.basename(file_path)}")
    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")
    captured_at = data.get("created_at") or data.get("captured_at")

    print(f"\n[Thông tin Capture]:")
    print(f" - App Name:     Coca-Cola Mini App")
    print(f" - App ID:       {COCA_APP_ID}")
    print(f" - Captured At:  {captured_at}")
    print(f" - Access Token: {access_token[:20]}...{access_token[-20:] if len(access_token) > 40 else ''}")
    print(f" - Refresh Token: {refresh_token[:20]}...{refresh_token[-20:] if len(refresh_token) > 40 else ''}")

    # Query Zalo Graph OpenAPI for user details (Standard Open API for Mini App token)
    print("\n[...] Đang thử lấy thông tin tài khoản Zalo thông qua Zalo OpenAPI...")
    zalo_api_url = "https://graph.zalo.me/v2.0/me?fields=id,name,picture"
    headers = {
        "access_token": access_token
    }
    
    try:
        response = requests.get(zalo_api_url, headers=headers, timeout=20)
        print("HTTP Status:", response.status_code)
        if response.status_code == 200:
            print("\n[OK] Thông tin tài khoản Zalo:")
            print(json.dumps(response.json(), indent=4, ensure_ascii=False))
        else:
            print("[WARN] Zalo OpenAPI trả về lỗi hoặc không dùng được token trực tiếp:")
            print(response.text)
    except Exception as e:
        print(f"[ERR] Lỗi kết nối Zalo OpenAPI: {e}")

    # Template cho API của chiến dịch Coca-Cola (AdtimaBox CRM / Campaign API)
    print("\n" + "-" * 60)
    print("LƯU Ý: Nếu Coca-Cola sử dụng CRM của AdtimaBox tương tự 7UP,")
    print("bạn cần cập nhật AUTH_URL và INFO_URL tương ứng để đổi token.")
    print("Ví dụ cấu trúc:")
    print(" - API thông tin: https://coca-coke-campaign-public-api.adtimabox.vn/digital-api/digital-user/me")
    print("-" * 60)

if __name__ == "__main__":
    main()
    input("\nNhấn Enter để kết thúc...")
