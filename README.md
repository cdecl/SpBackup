# SpBackup
SQL Server Stored Procedure Export

## Usage 
```
Usage: SpBackup.exe [connectionString] [path] [type]
  - connectionString : DB 연결문자열
  - path : 저장폴더(객체명으로 각각 생성) 혹은 저장파일(지정한 파일생성)
  - type : 객체타입
      ALL = ('FN', 'IF', 'P', 'TF', 'TR')
      P = 저장 프로시저 (default)
      FN = 스칼라 함수
      IF = 인라인 테이블 함수
      TF = 테이블 함수
      TR = 트리거
      V = 뷰
      U:<GetDDLProcName> = 테이블 (U:TableCreateScriptProcedure)
  - verion: 210202.1 (_MSC_VER:1928)
```
