TEAMPPT — A-1 "데이터 토대" 개발 실행 (Sonnet)

너는 이미 확정된 실행계획을 "정확히 그대로" 구현하는 개발 실행자다. 기획·설계는 Opus가
끝냈다. 방향을 새로 정하거나 범위를 임의로 넓히지 마라. 의심나면 멈추고 사용자에게 물어라.

[먼저 정독 — 이 순서로]
1. PROGRESS-BOARD.md (나라>대지>숲>나무>잎, 지금 위치 = A-1)
2. docs/superpowers/specs/2026-06-22-vibe-designing-a-first-design.md (이번 기획의 마스터.
   특히 §10 A-1 범위, §11 확정 결정 7개 + 유료화 진화경로)
3. 실행계획 3개 (이 순서로 실행):
   - docs/superpowers/plans/2026-06-22-a1a-llm-understanding-adapter.md   ← 지금 시작
   - docs/superpowers/plans/2026-06-22-a1bc-supabase-ingest-upload.md     ← Supabase 셋업 후
   - docs/superpowers/plans/2026-06-22-a1d-vector-read-path.md            ← b/c 후
4. 메모리: design_vibe_designing, project_teamppt, design_ai_redesign,
   feedback_coordinate_converter, feedback_admin_build, feedback_no_keys_in_docs,
   feedback_presentation

[실행 방식]
- superpowers:executing-plans (또는 subagent-driven-development)로 plan을 Task 단위로 실행.
- 각 plan의 스텝(- [ ])을 위에서 아래로. TDD 스텝(실패테스트→확인→구현→통과→커밋) 그대로 지켜라.
- 한 Task = 한 커밋. 커밋 메시지는 plan에 적힌 그대로 사용.
- **plan 경계를 넘을 때(A-1a 끝 → b/c 시작 등)는 자동 진행 금지. 사용자에게 보고하고 멈춰라.**

[지금 당장 할 것 = A-1a부터]
- A-1a는 Supabase 없이 가능(Gemini 키만 필요). 먼저 끝내라.
- A-1b/c·A-1d는 docs/SUPABASE-SETUP.md를 사용자가 완료해야 실행 가능 →
  착수 전 "Supabase 셋업(api-keys.json supabase 필드 + admin.json) 됐나요?" 확인하고 시작.

[건드리지 말 것 / 함정]
- Gemini 키는 이미 작동 확인됨(신형식, PPT 내 대화 검증). 키 형식이 'AIza'가 아니어도
  "잘못됐다"고 판단해 교체/수정하지 마라. (구 문서에 AQ.가 무효라는 기록이 있으나 해소됨.)
- 시크릿(키)을 문서·커밋에 평문으로 절대 넣지 마라. api-keys.json/admin.json은 gitignore.
- Core/ · Connect.cs · Globals.cs 직접 수정 금지. (유일 예외: A-1d Task 6의
  TaskPaneHost.cs 배선 — 그 plan에 적힌 범위만.)
- 의존성 추가 금지. Newtonsoft.Json + Office Interop + HttpClient만. supabase-csharp SDK 금지.
- CoordinateConverter에 폴백 로직 추가 금지(기존 non-fatal 패턴 유지).

[빌드/테스트]
- 일반 재컴파일(비관리자, COM 재등록 불필요):
  MSBuild TeampptAddin.csproj /t:Build /p:Configuration=Debug /p:Platform=AnyCPU /p:RegisterForComInterop=false
  MSBuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
- 단위테스트: 본프로젝트 빌드 → 테스트프로젝트 /p:BuildProjectReferences=false 빌드 →
  dotnet test --no-build --no-restore  (각 plan 맨 아래 "테스트 실행 절차" 참조)
- 관리자 빌드는 A-1d Task 6(UI 실구동 검증)에서만 필요. CLAUDE.md의 elevated RunAs cmd 래퍼는
  이 환경서 무력(build.log 갱신 안 됨) — 쓰지 마라.
- PowerPoint가 DLL 잠그므로 빌드 전 PPT 완전 종료.

[기록 — 매 plan 완료 시]
- PROGRESS-BOARD.md의 🌳나무/🍃잎을 현재 위치로 갱신(끝난 잎 교체, 골격 유지).
- docs/PITCH.md에 "완성된 기능"과 "버그 방지 설계 기능"을 비전문가 언어로 한 줄씩 누적
  (대표·투자자 발표 자산. 가치/안전/성장 프레임. 기술 자랑 금지). — memory feedback-presentation

[원칙 한 줄]
계획대로 정확히, 범위 임의확장 금지, plan 경계와 시크릿·핵심파일 앞에서 멈춰 확인.
