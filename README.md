# UnityExcelLoader

UnityExcelLoader는 Unity 에디터에서 Excel(.xls/.xlsx) 파일을 읽어 자동으로
- Entity 클래스(Serializable)와
- Script(ScriptableObject) 파일 및
- ScriptableObject 인스턴스(.asset)

를 생성해 주는 에디터 확장입니다.

이 README는 레포지토리의 실제 코드(특히 Editor 폴더의 구현)를 기반으로 작성되었습니다.

---

## 요약 / 핵심 동작
- 에디터 메뉴에서 엑셀 에셋을 선택한 뒤 다음 메뉴로 변환을 실행합니다:
  - Assets → Create → UnityExcelLoader → ExcelScript
  - Assets → Create → UnityExcelLoader → ExcelScriptableObject
- Excel 파일의 첫 행은 필드명(헤더), 두 번째 행은 타입(예: int, string, float, bool)을 기대합니다.
- 데이터 행은 3행(인덱스 2)부터 읽습니다.
- 헤더 이름 또는 행의 첫 번째 셀에 `#`이 있으면 해당 컬럼/레코드는 무시됩니다(주석 처리).
- 생성되는 파일 예시:
  - <ExcelName>_Entity.cs (엔티티 클래스)
  - <ExcelName>.cs (ScriptableObject 타입)
  - <ExcelName>.asset (데이터가 채워진 ScriptableObject 에셋)

---

## 요구사항 / 의존성
- Unity 2019.x 이상(패키지의 package.json에 `unity: "2019.1"` 명시)
- NPOI (엑셀 읽기용) — 코드에서 NPOI 네임스페이스를 사용합니다:
  - NPOI.HSSF.UserModel, NPOI.XSSF.UserModel, NPOI.SS.UserModel
- ScriptTemplate.txt, EntityTemplate.txt 템플릿 파일
  - 코드(Directory.GetFiles)에서 현재 작업 디렉터리(프로젝트 루트 이하)를 재귀 검색하므로 템플릿 파일은 프로젝트에 포함되어 있어야 합니다.
- (선택) Editor/Plugins 폴더에 DLL 배치(예: NPOI DLL)을 사용하거나 패키지 매니저로 추가

---

## 설치 방법
1. repository를 Packages/manifest.json에 git 의존성으로 추가하거나, 로컬으로 클론하여 패키지로 사용합니다. 예:
```json
"dependencies": {
  "com.piraxis.unity_excel_loader": "https://github.com/piraxis2/UnityExcelLoader.git"
}
```
또는 프로젝트에 직접 복사/임포트하세요.

---

## 사용법 (실제 레포지토리 동작 기준)
1. Unity 에디터에서 변환할 `.xls` 또는 `.xlsx` 파일을 Assets 폴더에 추가하고 선택합니다.
2. 우클릭 → Create → UnityExcelLoader → ExcelScript
   - 선택한 엑셀로부터 Entity(.cs)와 Script(.cs)를 생성합니다.
   - 스크립트 생성 시 기존 파일이 있으면 `//MethodsStart` / `//MethodsEnd` 사이의 내용을 보존합니다.
   - 동일하게 `//CustomUsingStart` / `//CustomUsingEnd` 영역의 using 지시문도 보존합니다.
3. 우클릭 → Create → UnityExcelLoader → ExcelScriptableObject
   - 이미 컴파일된 `<ExcelName>_Entity` 타입이 존재하는 경우(즉, 스크립트가 프로젝트에 있고 컴파일되었을 때) ScriptableObject 인스턴스(.asset)를 생성합니다.
   - 생성된 ScriptableObject의 `Data` 필드에 엑셀의 각 행이 entity 인스턴스로 담깁니다.
   - 저장 경로는 팝업으로 선택합니다(에디터에서 Save Folder Panel 호출).

Validation(메뉴 활성화 조건)
- ExcelScript 메뉴는 선택한 에셋이 Excel인지 체크합니다.
- ExcelScriptableObject 메뉴는 Excel 파일이면서 해당 엑셀명으로 된 엔티티 타입(예: MyTable_Entity)이 Assembly-CSharp에서 발견되는 경우에만 활성화됩니다.

---

## 엑셀 포맷 규칙 (구체)
- 헤더 행(0번째): 필드명(문자열). 예: Id, Name, Chance, IsUse
- 타입 행(1번째): 각 필드의 C# 타입명(예: int, string, float, bool). 비어 있으면 기본적으로 무시/처리 방식은 템플릿/코드에 의존.
- 데이터 행(2번째부터): 각 레코드 값
- 주석 처리:
  - 컬럼 헤더가 `#`로 시작하면 해당 컬럼은 생성에서 제외됩니다.
  - 데이터 행의 첫 번째 셀 값이 문자열이고 `#`로 시작하면 그 행 전체를 건너뜁니다.

값 변환:
- 에디터는 각 데이터 셀을 문자열로 취급하도록 셀 타입을 변경한 뒤, 필드 타입으로 Convert.ChangeType 시도를 합니다. 변환 불가 시 예외가 발생합니다(InvalidCastException, FormatException, OverflowException).

---

## 템플릿 파일 구조(중요)
코드는 ScriptTemplate.txt 와 EntityTemplate.txt를 읽어 생성물을 만듭니다. 템플릿에서 사용하는 플레이스홀더:
- {ClassName} — ScriptableObject 클래스 이름
- {Entity} — 엔티티 클래스 이름(예: {ExcelName}_Entity)
- {Fields} — 엔티티 필드 선언(엔티티 템플릿에서 사용)
- {Methods} — 보존된 커스텀 메소드 영역(기존 파일이 있을 경우 대체)
- {CustomUsing} — 보존된 using 선언 영역

템플릿에 `//MethodsStart` / `//MethodsEnd`, `//CustomUsingStart` / `//CustomUsingEnd` 같은 주석 영역을 만들어 두면, 재생성 시 사용자가 작성한 커스텀 코드가 유지됩니다.

예시
```csharp
//CustomUsingStart
{CustomUsing}
//CustomUsingEnd

[Serializable]
public class {ClassName}
{
{Fields}

    //MethodsStart
{Methods}
    //MethodsEnd
}
```

(위 구조는 예시이며 실제 템플릿 파일을 확인하세요.)

---

## 에디터 인스펙터
- UnityExcelLoader는 ScriptableObject 타입에 대해 커스텀 인스펙터(`UnityExcelLoaderEditor`)를 제공합니다.
- 인스펙터에서 `[ExcelScriptableObject]` 특성이 있는 ScriptableObject의 `data`(또는 `Data`) 프로퍼티를 스크롤 가능한 뷰(높이 300)로 보여줍니다.
- 표시되는 레이블: "Excel Scriptable Object"

---

## 생성 후 커스터마이즈 지침
- 자동 생성된 클래스 파일을 직접 편집할 때는 커스텀 메소드 영역(`//MethodsStart` / `//MethodsEnd`)과 커스텀 using 영역(`//CustomUsingStart` / `//CustomUsingEnd`)에 코드를 추가하세요. 재생성 시 이 영역은 유지됩니다.
- ScriptableObject의 데이터에 접근할 때는 생성된 `<ExcelName>` 클래스의 `Data` 필드를 사용하세요.

---

## 문제해결(팁)
- 메뉴가 비활성화되거나 타입을 찾을 수 없다는 메시지가 뜨면
  - 먼저 스크립트가 컴파일되어 `<ExcelName>_Entity` 타입이 Assembly-CSharp에서 노출되는지 확인하세요.
  - NPOI가 프로젝트에 포함되어 컴파일되는지 확인하세요.
  - 템플릿 파일의 존재 위치를 확인하세요(프로젝트 루트 또는 검색 가능한 위치).
- 타입 변환 예외 발생 시
  - 엑셀의 타입 행(2번째 행)에 지정한 타입과 실제 데이터가 일치하는지 확인하세요.
  - 숫자/소수점/불리언 형식 등을 엑셀 셀에서 올바르게 포맷했는지 확인하세요.

---

## 예제
(간단한 엑셀 구조 예)
| Id | Name  | Chance | IsUse |
|----|-------|--------|-------|
| int| string| float  | bool  |
| 1  | Apple | 0.5    | TRUE  |
| 2  | Banana| 0.2    | FALSE  |

생성 결과:
- MySheet_Entity.cs
- MySheet.cs (ScriptableObject 타입, List<MySheet_Entity> Data 포함)
- MySheet.asset (해당 데이터가 담긴 에셋)
