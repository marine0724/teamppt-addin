# TEAMPPT Add-in — 개발 가이드

## 빌드

- **MSBuild 경로**: `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`
- **COM 등록 때문에 관리자 권한 필수** (`RegisterForComInterop=true`)
- 빌드 명령 (관리자 권한):

```powershell
Start-Process -FilePath "cmd.exe" -ArgumentList '/c "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\Projects\teamppt-addin\src\TeampptAddin\TeampptAddin.sln" /t:Build /p:Configuration=Debug "/p:Platform=Any CPU" /verbosity:minimal > "c:\Projects\teamppt-addin\build.log" 2>&1' -Verb RunAs -Wait -WindowStyle Hidden
```

- 빌드 결과는 `build.log` 파일 끝 5줄 확인 (`tail -5 build.log`)
- `/t:Build` 사용 (변경분만 빌드). 전체 재빌드 필요시 `/t:Rebuild`

## API 키

- API 키를 문서나 커밋에 평문으로 절대 포함하지 않는다
