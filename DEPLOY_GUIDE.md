# GitHub + Netlify 배포 가이드

## 1단계: GitHub 리포 만들기

1. https://github.com/new 에서 새 리포 생성
   - Repository name: `hwpx-ai-generator` (원하는 이름)
   - **Private** 선택 (API 키 보안)
   - README 체크 해제
   - Create repository 클릭

## 2단계: 로컬에서 Push

PowerShell 또는 CMD에서 이 폴더로 이동 후 실행:

```powershell
cd C:\Users\value\Desktop\_claudecowork\hwpx-browser-demo

# git 초기화 (이미 .git 폴더가 있으면 삭제 후)
rmdir /s /q .git
git init -b main
git add .gitignore netlify.toml index.html demo.html generate_hwpx_browser.js generate_hwpx_node.js hwpx_template_browser.js hwpx_template_node.js hwpx_press_template.js netlify/functions/chat.mjs
git commit -m "HWPX AI 문서 생성기 - 서면보고 + 보도자료 멀티 템플릿"

# GitHub 리포 연결 (아래 URL을 본인 리포로 변경)
git remote add origin https://github.com/YOUR_USERNAME/hwpx-ai-generator.git
git push -u origin main
```

## 3단계: Netlify 연동

1. https://app.netlify.com 접속
2. **Add new site** → **Import an existing project**
3. **GitHub** 선택 → `hwpx-ai-generator` 리포 선택
4. Build settings:
   - Build command: (비워두기)
   - Publish directory: `.`
5. **Deploy site** 클릭

## 4단계: 환경변수 설정

1. Netlify 대시보드 → **Site configuration** → **Environment variables**
2. **Add a variable** 클릭
3. Key: `CLAUDE_API_KEY`, Value: 본인 API 키 입력
4. **Save** 클릭

## 5단계: 재배포

환경변수 설정 후 **Deploys** 탭 → **Trigger deploy** → **Deploy site**

## 이후 업데이트

코드 수정 후:
```powershell
git add -A
git commit -m "변경 내용 설명"
git push
```
→ Netlify가 자동으로 재배포합니다.
