# talk-deck 스킬 v1.1 (2026-11-12)

데이터·결과가 이미 있는 내용을 발표용 PPT로 만드는 에이전트 스킬이에요. 기업 대상 데이터·제품 설명, 행사 발표, 연구 발표에 쓸 수 있어요.
Codex와 Claude Code가 같은 형식(SKILL.md)을 읽기 때문에 폴더 하나로 양쪽에서 써요.

## 들어 있는 것

| 파일 | 내용 |
|---|---|
| `talk-deck/SKILL.md` | 작업 순서: 질문 → 스토리라인 → 사실 확인표 → 3관점 검토 → 자료 모으기 → 빌드 → 렌더·눈 검수 → 수정 라운드 → 최종 검토 |
| `talk-deck/references/collecting_materials.md` | 사용자에게 요구할 것과 에이전트가 수집할 것, 사진 라이선스 확인 |
| `talk-deck/references/storyline_and_review.md` | 한 줄 메시지, 발표 종류별 장 구성, 검토 지시문, 대본 쓰는 법 |
| `talk-deck/references/design_system.md` | "AI가 만든 슬라이드" 증상 11가지, 글자 4단계, 색, 슬라이드 틀 |
| `talk-deck/references/final_review_checklist.md` | 주장 강도, 일치 검사, 디자인·기술 검사 |
| `talk-deck/scripts/` | `check_env.py`(환경 점검·패키지 설치), `pptx_kit.py`(python-pptx 헬퍼), `render_pptx.ps1`(PNG 렌더), `overview_grid.py`(전체 썸네일), `export_slides.py`(일부 장만 저장) |

## 필요한 환경

| 항목 | 필요 여부 | 메모 |
|---|---|---|
| 파일을 읽고 스크립트를 실행하는 에이전트 (Codex CLI·IDE, Claude Code) | 필수 | 채팅창에 붙여 넣는 방식으로는 렌더·검수 단계가 돌지 않아요 |
| Python 3 | 필수 | 직접 설치해야 해요 |
| `python-pptx`, `Pillow`, `matplotlib` | 필수(matplotlib은 선택) | 에이전트가 처음 쓸 때 `check_env.py --install`로 설치해요 |
| Windows + PowerPoint | 권장 | 렌더가 실제 PowerPoint 화면과 같아요 |
| LibreOffice + poppler(`pdftoppm`) | PowerPoint가 없을 때 | Mac·Linux용 대체 경로. 글꼴·줄바꿈이 조금 달라요 |
| 웹 검색 도구 | 권장 | 없으면 배경 자료·사진을 에이전트가 수집하지 못하고 목록만 넘겨요 |

## 설치

스킬만 받으면 돼요. 다른 스킬은 필요 없고, 빠진 Python 패키지는 에이전트가 처음 실행할 때 점검해서 설치해요. PowerPoint나 LibreOffice 같은 프로그램은 자동으로 설치하지 않고 먼저 물어봐요.

가장 쉬운 방법은 저장소를 스킬 폴더 아래에 바로 받는 거예요.

```
# Codex (Windows 명령 프롬프트)
git clone https://github.com/nicekwmin-gif/talk-deck-skill "%USERPROFILE%\.agents\skills\talk-deck-skill"

# Codex (Mac·Linux·PowerShell)
git clone https://github.com/nicekwmin-gif/talk-deck-skill ~/.agents/skills/talk-deck-skill

# Claude Code
git clone https://github.com/nicekwmin-gif/talk-deck-skill ~/.claude/skills/talk-deck-skill
```

새 버전은 그 폴더에서 `git pull`로 받아요. 스킬이 목록에 안 보이면 저장소 안의 `talk-deck` 폴더만 아래 표의 위치로 옮겨요.

직접 복사하려면 `talk-deck` 폴더를 통째로 아래 위치에 둬요.

| 도구 | 개인용 (모든 프로젝트) | 프로젝트용 |
|---|---|---|
| Codex | `~/.agents/skills/talk-deck/` (Windows: `C:\Users\<이름>\.agents\skills\talk-deck\`) | `<프로젝트>/.agents/skills/talk-deck/` |
| Claude Code | `~/.claude/skills/talk-deck/` | `<프로젝트>/.claude/skills/talk-deck/` |

- Codex 경로는 OpenAI 공식 문서(Build skills, 2026-11-12 확인) 기준이에요. 스킬이 안 보이면 Codex를 다시 시작해요.
- Codex에서는 `/skills`로 목록을 보거나 `$talk-deck`으로 직접 불러요. Claude Code에서는 `/talk-deck`이에요.

## 쓰는 법

1. 발표용 폴더를 하나 만들고 가진 자료(데이터, 그림, 사진, 이전 발표 자료, 로고)를 넣어요.
2. 그 폴더에서 에이전트를 열고 이렇게 시작해요.

   > `$talk-deck` 이 폴더 자료로 OO사 대상 15분 발표 자료를 만들어 줘. 청중은 OO팀 실무자와 임원이야.

3. 에이전트가 필요한 것을 한 번에 물어요. 답하고, 스토리라인(글)부터 확인해요.
4. 슬라이드가 나오면 PowerPoint에서 열어 보고 고칠 곳을 말해요. "3장 그래프가 작다", "5장은 빼자"처럼 장 번호로 말하면 돼요.
5. 직접 고친 pptx를 다시 주면 그 파일을 기준으로 이어서 고쳐요.

## 알아 둘 것

- **한 번에 완성본이 나오지 않아요.** 이 방식의 품질은 스토리라인 확인과 수정 라운드에서 나와요. 원래 덱(15분 발표)도 사흘 동안 수십 번 고쳐서 만들었어요.
- 수치·인용은 받은 자료와 직접 확인한 출처에서만 써요. 확인 못 한 것은 `[확인]`으로 표시돼요.
- 공개하면 안 되는 수치·고객사 이름은 처음에 알려 줘야 해요.
- 회사 템플릿(pptx/potx)이 있으면 같이 넣어 주세요. 그 글꼴·색을 따라요.

## 검증한 범위

| 항목 | 상태 |
|---|---|
| `pptx_kit.py` 빌드, `render_pptx.ps1` 렌더, `overview_grid.py`, `export_slides.py` | Windows + PowerPoint + Python 3.13에서 3장짜리 시험 덱으로 확인 (2026-11-12) |
| Codex에서 스킬 인식·실행 | 확인 못 함. 이 PC에는 Codex가 없어요 |
| LibreOffice 렌더 경로 | 확인 못 함 |
