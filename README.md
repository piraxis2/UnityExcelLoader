
# Language
[en](https://github.com/piraxis2/UnityExcelLoader/tree/main?tab=readme-ov-file#unityexcelloader_en)
[kr](https://github.com/piraxis2/UnityExcelLoader/tree/main?tab=readme-ov-file#unityexcelloader_kr)

# UnityExcelLoader

UnityExcelLoader is a tool to convert Excel files into ScriptableObjects in Unity projects.

## Installation

To install this package, add the following dependency to your `Packages/manifest.json` file:

```json
{
  "dependencies": {
    "com.piraxis2.unityexcelloader": "https://github.com/piraxis2/UnityExcelLoader.git"
  }
}
```

## Usage

### Convert Excel file to ScriptableObject

1. Add the Excel file to your Unity project.
2. Right-click the Excel file and select `Assets/Create/UnityExcelLoader/ExcelScript`.
3. Choose the folder to save, and the corresponding ScriptableObject and Entity class will be generated.

### Use the generated ScriptableObject

1. You can inspect the generated ScriptableObject in the Inspector.
2. Use the data in the ScriptableObject to implement game logic.

### Add custom methods

The generated ScriptableObject and Entity classes have a custom methods section. This section is located between the `//MethodsStart` and `//MethodsEnd` comments. You can add methods between these comments. For example, to add a custom method to the `Cube_Grade` class:

```csharp
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityExcelLoader.Runtime;

namespace UnityExcelLoader
{
    [ExcelScriptableObject]
    public class Cube_Grade : ScriptableObject
    {
        public List<Cube_Grade_Entity> Data;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        public void PrintData()
        {
            foreach (var entity in Data)
            {
                Debug.Log($"Id: {entity.Id}, Grade: {entity.Grade}");
            }
        }
        //MethodsEnd
    }
}
```

## Example

### Cube_Grade ScriptableObject

`Cube_Grade` is an example of a ScriptableObject generated from an Excel file.

```csharp
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityExcelLoader.Runtime;

namespace UnityExcelLoader
{
    [ExcelScriptableObject]
    public class Cube_Grade : ScriptableObject
    {
        public List<Cube_Grade_Entity> Data;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        //MethodsEnd
    }
}
```

### Cube_Grade_Entity class

`Cube_Grade_Entity` is the data entity class for the `Cube_Grade` ScriptableObject.

```csharp
using System;
using System.Collections.Generic;
using UnityEngine;

namespace UnityExcelLoader
{
    [Serializable]
    public class Cube_Grade_Entity
    {
        public int Id;
        public int Grade;
        public string Unit_String;
        public string Value;
        public string Cube_Color;
        public string Txt_Color;
        public bool Cube_Create;
        public bool Cube_Add_Delete;
        public int Cube_Min_Id;
        public int Cube_Max_Id;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        //MethodsEnd
    }
}
```

## Contribution

If you want to contribute, please fork this repository and send a pull request.

## License

This project is distributed under the MIT License. See the `LICENSE` file for more details.









# UnityExcelLoader

UnityExcelLoader는 Unity 프로젝트에서 Excel 파일을 ScriptableObject로 변환하는 도구입니다.

## 설치

이 패키지를 설치하려면 `Packages/manifest.json` 파일에 다음 의존성을 추가하세요:

```json
{
  "dependencies": {
    "com.piraxis2.unityexcelloader": "https://github.com/piraxis2/UnityExcelLoader.git"
  }
}
```

## 사용법

### 엑셀 파일을 ScriptableObject로 변환

1. Unity 에디터에서 엑셀 파일을 프로젝트에 추가합니다.
2. 엑셀 파일을 우클릭하고 `Assets/Create/UnityExcelLoader/ExcelScript` 메뉴를 선택합니다.
3. 저장할 폴더를 선택하면, 엑셀 파일에 대응하는 ScriptableObject와 Entity 클래스가 생성됩니다.

### 생성된 ScriptableObject 사용

1. 생성된 ScriptableObject를 인스펙터에서 확인할 수 있습니다.
2. ScriptableObject의 데이터를 사용하여 게임 로직을 구현할 수 있습니다.

### 커스텀 메소드 추가

생성된 ScriptableObject와 Entity 클래스에는 개발자가 사용할 수 있는 커스텀 메소드 영역이 있습니다. 이 영역은 `//MethodsStart`와 `//MethodsEnd` 주석 사이에 위치하며, 이 주석 사이에 메소드를 추가할 수 있습니다. 예를 들어, `Cube_Grade` 클래스에 커스텀 메소드를 추가하려면 다음과 같이 작성할 수 있습니다:

```csharp
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityExcelLoader.Runtime;

namespace UnityExcelLoader
{
    [ExcelScriptableObject]
    public class Cube_Grade : ScriptableObject
    {
        public List<Cube_Grade_Entity> Data;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        public void PrintData()
        {
            foreach (var entity in Data)
            {
                Debug.Log($"Id: {entity.Id}, Grade: {entity.Grade}");
            }
        }
        //MethodsEnd
    }
}
```

## 예제

### Cube_Grade ScriptableObject

`Cube_Grade`는 Excel 파일에서 생성된 ScriptableObject의 예제입니다.

```csharp
using System;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityExcelLoader.Runtime;

namespace UnityExcelLoader
{
    [ExcelScriptableObject]
    public class Cube_Grade : ScriptableObject
    {
        public List<Cube_Grade_Entity> Data;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        //MethodsEnd
    }
}
```

### Cube_Grade_Entity 클래스

`Cube_Grade_Entity`는 `Cube_Grade` ScriptableObject의 데이터 엔티티 클래스입니다.

```csharp
using System;
using System.Collections.Generic;
using UnityEngine;

namespace UnityExcelLoader
{
    [Serializable]
    public class Cube_Grade_Entity
    {
        public int Id;
        public int Grade;
        public string Unit_String;
        public string Value;
        public string Cube_Color;
        public string Txt_Color;
        public bool Cube_Create;
        public bool Cube_Add_Delete;
        public int Cube_Min_Id;
        public int Cube_Max_Id;

        //DoNotRemove Methods Start End Annotation
        //MethodsStart
        //MethodsEnd
    }
}
```

## 기여

기여를 원하시면, 이 저장소를 포크하고 풀 리퀘스트를 보내주세요.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.
