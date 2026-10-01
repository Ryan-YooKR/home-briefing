"""
github_publisher.py
GitHub Pages에 주간 부동산 리포트 HTML을 자동 배포/업데이트하는 모듈
Git CLI 설치 없이 GitHub REST API(HTTPS)만으로 100% 자동 동작합니다.
"""

import os
import sys
import json
import base64
import requests

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "github_config.json")

def load_github_config():
    """github_config.json 파일 또는 환경 변수(GitHub Actions)에서 설정 불러오기"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"❌ [오류] github_config.json 읽기 실패: {e}")
    
    # 환경 변수 지원 (GitHub Actions 등)
    gh_token = os.environ.get("GH_PAT") or os.environ.get("GITHUB_TOKEN")
    repo_env = os.environ.get("GITHUB_REPOSITORY", "Ryan-YooKR/home-briefing")
    if gh_token and "/" in repo_env:
        owner, repo = repo_env.split("/", 1)
        return {
            "owner": owner,
            "repo": repo,
            "token": gh_token,
            "pages_url": f"https://{owner.lower()}.github.io/{repo}/"
        }
    return None

def save_github_config(owner, repo, token):
    """github_config.json 파일에 설정 저장"""
    config = {
        "owner": owner.strip(),
        "repo": repo.strip(),
        "token": token.strip(),
        "pages_url": f"https://{owner.strip().lower()}.github.io/{repo.strip()}/"
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
    print(f"✅ [설정 저장] GitHub 연동 정보가 {CONFIG_FILE} 에 안전하게 저장되었습니다.")
    return config

def publish_html_to_github_pages(html_content, report_date=""):
    """
    GitHub 저장소의 index.html 파일을 업데이트하여 GitHub Pages로 자동 배포
    """
    config = load_github_config()
    if not config:
        print("⚠️ [안내] github_config.json 설정이 없습니다. GitHub 연동을 먼저 진행해 주세요.")
        return None

    owner = config.get("owner")
    repo = config.get("repo")
    token = config.get("token")
    pages_url = config.get("pages_url", f"https://{owner.lower()}.github.io/{repo}/")

    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "HosangSunkyung-RealEstateReporter"
    }

    api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/index.html"

    # 1. 기존 index.html의 SHA 확인 (업데이트 시 필요)
    sha = None
    try:
        r_get = requests.get(api_url, headers=headers, timeout=10)
        if r_get.status_code == 200:
            sha = r_get.json().get("sha")
        elif r_get.status_code == 404:
            # 파일이 아직 없음 (신규 생성)
            pass
        elif r_get.status_code == 401:
            print("❌ [오류] GitHub 토큰(Token) 인증에 실패했습니다. 토큰을 다시 확인해 주세요.")
            return None
    except Exception as e:
        print(f"❌ [오류] GitHub API 통신 실패: {e}")
        return None

    # 2. HTML 콘텐츠 Base64 인코딩
    content_bytes = html_content.encode("utf-8")
    content_b64 = base64.b64encode(content_bytes).decode("utf-8")

    commit_message = f"주간 부동산 리포트 자동 배포 ({report_date})" if report_date else "주간 부동산 리포트 업데이트"

    payload = {
        "message": commit_message,
        "content": content_b64
    }
    if sha:
        payload["sha"] = sha

    # 3. PUT 요청으로 index.html 생성 또는 업데이트
    try:
        r_put = requests.put(api_url, headers=headers, json=payload, timeout=15)
        if r_put.status_code in [200, 201]:
            print(f"🎉 [성공] GitHub Pages(index.html) 배포 완료!")
            print(f"🌐 [모바일 열람 웹 링크] {pages_url}")
            return pages_url
        else:
            print(f"❌ [배포 실패] GitHub API 오류 ({r_put.status_code}): {r_put.text}")
            return None
    except Exception as e:
        print(f"❌ [오류] 파일 업로드 중 예외 발생: {e}")
        return None

if __name__ == "__main__":
    cfg = load_github_config()
    if cfg:
        print("현재 설정된 GitHub Pages 주소:", cfg.get("pages_url"))
    else:
        print("GitHub 연동 정보가 아직 설정되지 않았습니다.")
