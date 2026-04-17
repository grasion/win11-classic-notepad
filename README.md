# Win11 메모장 버전 전환 도구

윈도우 11의 느린 신버전 메모장과 빠른 구버전(클래식) 메모장을 **원클릭으로 전환**하는 도구입니다.

## 기능

- **자동 감지** — 실행 시 현재 신버전/구버전 상태를 자동으로 판별
- **양방향 전환** — 버튼 하나로 구버전 ↔ 신버전 전환
- **구버전 전환 시 수행 작업:**
  - 레지스트리 `NoOpenWith` 값 삭제
  - 신버전 메모장 앱(Microsoft.WindowsNotepad) 제거
  - `.txt`, `.log`, `.ini`, `.cfg`, `.inf` 파일 연결 설정
  - 작업표시줄에 클래식 메모장 고정
- **신버전 복원 시 수행 작업:**
  - `NoOpenWith` 레지스트리 값 복원
  - winget을 통한 신버전 메모장 재설치 (실패 시 Microsoft Store 자동 열기)

## 스크린샷

> <img width="621" height="638" alt="image" src="https://github.com/user-attachments/assets/ca29f8f4-36f1-4bd7-983a-563dc4b8303b" />


## 사용법

1. [Releases](../../releases) 페이지에서 `ClassicNotepadRestore.exe` 다운로드
2. 더블클릭하여 실행 (관리자 권한 자동 요청)
3. 현재 상태를 확인하고 전환 버튼 클릭

## 빌드 방법

.NET 8.0 SDK 이상이 필요합니다.

```bash
dotnet publish ClassicNotepadRestore/ClassicNotepadRestore.csproj ^
  -c Release -r win-x64 --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:IncludeNativeLibrariesForSelfExtract=true ^
  -o ./publish
```

## 참고

- [윈도우11 메모장 구버전으로 사용하기 (jabda-blog)](https://jabda-blog.tistory.com/177)
- 윈도우 대규모 업데이트(24H2 등) 후 `NoOpenWith`가 다시 생길 수 있습니다. 그때 다시 실행하세요.

## 라이선스

MIT License
