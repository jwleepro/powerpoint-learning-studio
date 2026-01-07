# PowerPoint Learning Studio 코드 품질 개선 계획

## 1. 현재 상태 분석 요약

### 발견된 주요 문제점

| 문제 | 심각도 | 위치 |
|-----|--------|------|
| 매직 넘버 (10+ 곳) | 높음 | PowerPointService.cs |
| 테스트 중복 코드 (6회) | 높음 | PowerPointElementPropertyTests.cs |
| 모니터링 패턴 반복 (3세트) | 중간 | PowerPointService.cs |
| 하드코딩 문자열 (10+ 곳) | 중간 | 여러 파일 |
| SOLID 원칙 위반 | 중간 | PowerPointService.cs |
| 빈 플레이스홀더 파일 | 낮음 | Class1.cs, UnitTest1.cs |

### SOLID 원칙 위반 상세

- **SRP 위반**: PowerPointService가 7가지 책임 보유 (621줄, 34개 메서드)
- **OCP 위반**: 확장 시 기존 클래스 수정 필요
- **LSP 위반**: 인터페이스 부재
- **ISP 위반**: 모든 기능이 단일 클래스에 통합
- **DIP 위반**: 구체 클래스에 직접 의존

---

## 2. 개선 계획 (TDD Tidy First 방식)

### Phase 1: 빈 파일 제거

**작업**: 사용하지 않는 빈 파일 삭제

| 삭제 파일 |
|-----------|
| `src/main/PPTCoach.Core/Class1.cs` |
| `src/test/PPTCoach.Tests/UnitTest1.cs` |

---

### Phase 2: 상수 추출

**작업**: 매직 넘버와 하드코딩 문자열을 상수로 추출

#### 2.1 PowerPoint COM 상수 클래스 생성

**신규 파일**: `src/main/PPTCoach.Core/Constants/PowerPointConstants.cs`

```csharp
namespace PPTCoach.Core.Constants;

public static class PpSlideLayout
{
    public const int Blank = 12;
    public const int Title = 1;
    public const int Text = 2;
}

public static class PpViewType
{
    public const int Normal = 1;
    public const int SlideSorter = 3;
}

public static class PpSelectionType
{
    public const int None = 0;
    public const int Slides = 1;
    public const int Shapes = 2;
}

public static class MsoTriState
{
    public const int True = -1;
    public const int False = 0;
}

public static class MsoShapeType
{
    public const int AutoShape = 1;
    public const int Picture = 13;
    public const int Placeholder = 14;
    public const int TextBox = 17;
    public const int Table = 19;
}

public static class PowerPointProgId
{
    public const string Application = "PowerPoint.Application";
    public const string Executable = "powerpnt.exe";
}
```

#### 2.2 수정 대상

| 파일 | 변경 내용 |
|------|----------|
| `PowerPointService.cs` | 매직 넘버 → 상수 참조 |
| `PowerPointTestHelpers.cs` | "powerpnt.exe" → 상수 참조 |

---

### Phase 3: 테스트 중복 제거

**작업**: 인라인 PowerPoint 시작 로직을 헬퍼 메서드 호출로 교체

**수정 파일**: `src/test/PPTCoach.Tests/Phase01/PowerPointElementPropertyTests.cs`

**변경 메서드** (6개):
- `ShouldGetFontNameFromTextShape`
- `ShouldGetFontSizeFromTextShape`
- `ShouldGetFontColorFromTextShape`
- `ShouldGetShapeFillColor`
- `ShouldGetTableCellContent`
- `ShouldHandleShapesWithoutTextGracefully`

**변경 전**:
```csharp
var psi = new ProcessStartInfo { FileName = "powerpnt.exe", ... };
pptProcess = Process.Start(psi);
const int maxRetries = 20;
const int retryDelayMs = 500;
for (int i = 0; i < maxRetries; i++) { ... }
```

**변경 후**:
```csharp
(pptProcess, instance) = PowerPointTestHelpers.StartPowerPointAndWait(
    powerPointService, ProcessWindowStyle.Minimized);
```

---

### Phase 4: 인터페이스 분리 (ISP/DIP)

**작업**: 책임별 인터페이스 추출

#### 4.1 폴더 구조

```
src/main/PPTCoach.Core/
├── Interfaces/
│   ├── IPowerPointConnection.cs
│   ├── IPresentationManager.cs
│   ├── ISlideOperations.cs
│   ├── ISlideQuery.cs
│   ├── IShapeQuery.cs
│   ├── IFontQuery.cs
│   ├── IPresentationMonitor.cs
│   └── IPowerPointService.cs (파사드)
```

#### 4.2 인터페이스 정의

| 인터페이스 | 메서드 |
|------------|--------|
| `IPowerPointConnection` | IsPowerPointInstalled, GetRunningPowerPointInstance, ConnectToPowerPointOrThrow |
| `IPresentationManager` | CreateNewPresentation |
| `ISlideOperations` | AddBlankSlide, AddSlideWithLayout, DeleteSlideByIndex, MoveSlide |
| `ISlideQuery` | GetCurrentSlideNumber, GetTotalSlideCount, GetSlideTitle, GetAllTextFromSlide, GetShapesOnSlide |
| `IShapeQuery` | GetShapeType, GetShapeFillColor, GetShapePosition, GetShapeSize, GetTableCellContent |
| `IFontQuery` | GetFontName, GetFontSize, GetFontColor |
| `IPresentationMonitor` | Start/Stop/CheckFor* (슬라이드, Shape선택, 저장) + 3개 이벤트 |
| `IPowerPointService` | 위 모든 인터페이스 상속 (파사드) |

---

### Phase 5: 서비스 클래스 분리 (SRP)

**작업**: 책임별 구현 클래스 추출

#### 5.1 폴더 구조

```
src/main/PPTCoach.Core/
├── Services/
│   ├── PowerPointConnectionService.cs
│   ├── PresentationManagerService.cs
│   ├── SlideOperationsService.cs
│   ├── SlideQueryService.cs
│   ├── ShapeQueryService.cs
│   ├── FontQueryService.cs
│   └── PresentationMonitorService.cs
├── Utilities/
│   └── ComRgbConverter.cs
```

#### 5.2 분리 순서 (의존성 고려)

| 순서 | 서비스 | 의존성 |
|------|--------|--------|
| 1 | ComRgbConverter | 없음 |
| 2 | PowerPointConnectionService | 없음 |
| 3 | PresentationManagerService | 없음 |
| 4 | SlideOperationsService | 없음 |
| 5 | SlideQueryService | 없음 |
| 6 | ShapeQueryService | ComRgbConverter |
| 7 | FontQueryService | ComRgbConverter |
| 8 | PresentationMonitorService | ISlideQuery (DIP) |

---

### Phase 6: 파사드 패턴 적용 (기존 API 호환성)

**작업**: PowerPointService를 파사드로 변환

**파일**: `src/main/PPTCoach.Core/PowerPointService.cs`

```csharp
public class PowerPointService : IPowerPointService
{
    private readonly IPowerPointConnection _connection;
    private readonly IPresentationManager _presentationManager;
    private readonly ISlideOperations _slideOperations;
    private readonly ISlideQuery _slideQuery;
    private readonly IShapeQuery _shapeQuery;
    private readonly IFontQuery _fontQuery;
    private readonly IPresentationMonitor _monitor;

    // 기본 생성자 (기존 호환성)
    public PowerPointService()
    {
        _connection = new PowerPointConnectionService();
        _presentationManager = new PresentationManagerService();
        // ...
    }

    // DI 생성자 (테스트/확장용)
    public PowerPointService(
        IPowerPointConnection connection,
        IPresentationManager presentationManager,
        // ...
    ) { }

    // 모든 메서드는 하위 서비스에 위임
    public bool IsPowerPointInstalled() => _connection.IsPowerPointInstalled();
    // ...
}
```

---

## 3. 실행 순서 및 검증

| 순서 | Phase | 검증 방법 |
|------|-------|----------|
| 1 | Phase 1: 빈 파일 제거 | `dotnet build` |
| 2 | Phase 2: 상수 추출 | `dotnet build && dotnet test` |
| 3 | Phase 3: 테스트 중복 제거 | `dotnet test` |
| 4 | Phase 4: 인터페이스 분리 | `dotnet build && dotnet test` |
| 5 | Phase 5: 서비스 분리 | `dotnet build && dotnet test` (각 서비스별) |
| 6 | Phase 6: 파사드 완성 | `dotnet build && dotnet test` |

---

## 4. 수정 대상 파일 요약

### 신규 생성 (17개)

```
src/main/PPTCoach.Core/
├── Constants/
│   └── PowerPointConstants.cs
├── Interfaces/
│   ├── IPowerPointConnection.cs
│   ├── IPresentationManager.cs
│   ├── ISlideOperations.cs
│   ├── ISlideQuery.cs
│   ├── IShapeQuery.cs
│   ├── IFontQuery.cs
│   ├── IPresentationMonitor.cs
│   └── IPowerPointService.cs
├── Services/
│   ├── PowerPointConnectionService.cs
│   ├── PresentationManagerService.cs
│   ├── SlideOperationsService.cs
│   ├── SlideQueryService.cs
│   ├── ShapeQueryService.cs
│   ├── FontQueryService.cs
│   └── PresentationMonitorService.cs
└── Utilities/
    └── ComRgbConverter.cs
```

### 수정 (3개)

| 파일 | 변경 내용 |
|------|----------|
| `PowerPointService.cs` | 파사드로 변환, 하위 서비스에 위임 |
| `PowerPointTestHelpers.cs` | 상수 참조 |
| `PowerPointElementPropertyTests.cs` | 중복 코드 제거 |

### 삭제 (2개)

| 파일 |
|------|
| `src/main/PPTCoach.Core/Class1.cs` |
| `src/test/PPTCoach.Tests/UnitTest1.cs` |

---

## 5. SOLID 원칙 적용 결과

| 원칙 | 적용 전 | 적용 후 |
|------|--------|--------|
| **SRP** | 1클래스 7책임 | 7클래스 각 1책임 |
| **OCP** | 확장=수정 | 인터페이스로 확장 가능 |
| **LSP** | 인터페이스 없음 | 명확한 계약 정의 |
| **ISP** | 34메서드 통합 | 7개 세분화 인터페이스 |
| **DIP** | 구체 클래스 의존 | 인터페이스 주입 |
