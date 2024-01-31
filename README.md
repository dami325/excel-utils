# 소개
엑셀 다운로드 기능의 생산성과 보일러 플레이트 코드 제거를 위해 만든 라이브러리입니다.<br>
객체 필드에 어노테이션을 추가하는 것만으로 간단하게 엑셀 다운로드를 할 수 있습니다.<br>
[어노테이션 속성](https://github.com/dami325/excel-utils/blob/master/src/main/java/net/youyoung/excel/annotation/ExcelColumn.java)에는 아래와 같은 기능이 있습니다.
1. header - 필드의 헤더명을 지정합니다.
2. headerEn - 사용자 브라우저 언어 정보로 한국이 아닐때 헤더명을 적용합니다. (미입력 시 header 값)
3. width - 셀의 너비를 지정할 수 있습니다.
4. headerStyle - 헤더의 스타일을 지정할 수 있습니다.
5. bodyStyle - 필드의 스타일을 지정할 수 있습니다.
6. format - 날짜 필드, 숫자 필드 등 셀 포멧 형식을 지원합니다. ("#.###", "yyyy-MM-dd" 등)
7. columnDefault - 필드 값이 없을 경우 셀에 보여줄 필드의 기본값을 지정할 수 있습니다. ("-" 등)
<br>

[여기](https://techblog.woowahan.com/2698/)가 많은 도움이 되었습니다.<br>

<br>

# 사용방법

### 1. 의존성을 추가합니다.
[메이븐 저장소](https://mvnrepository.com/artifact/io.github.dami325/excel-utils)
#### gradle
```
implementation group: 'io.github.dami325', name: 'excel-utils', version: '0.0.2'
```
#### maven
```
<dependency>
    <groupId>io.github.dami325</groupId>
    <artifactId>excel-utils</artifactId>
    <version>0.0.2</version>
</dependency>
```
### 2. 다운받을 리스트 객체 필드에 어노테이션 정보를 추가합니다.
```
public class ExcelDownloadExample {

    @ExcelColumn(header = "이름")
    private String name;

    ...
}
```
### 3. ExcelUtils.download() 를 호출합니다.
```
public void downloadExcel() {
    List<ExcelDownloadExample> list = exampleRepository.findAll();
    ExcelUtils.download(list,ExcelDownloadExample.class, "다운받을 파일 이름");
}
```
