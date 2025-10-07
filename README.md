# LifeTennisMatcher

## 기능
- 참가자 명단 읽기
- 혼합 복식 시퀀스 결정
- 라이프 회원 그룹 정보 로드 (데이터 저장)
- 매칭 로직 (혼복, 남복, 여복)
- 매칭 중복 방지를 위한 스왑 로직
- 매칭 결과 엑셀 저장
- 파일명 자동 증가 저장
- 라이프 멤버일 경우 이름에 * 추가 (동명이인 게스트의 경우 이름 뒤에 번호를 추가하여 구분)
- 성별 식별자 추가 (남자: (m), 여자: (f))

## 사용 방법
1. `Auto_Table.xlsx` 파일에 참가자 명단을 입력합니다.
2. 스크립트를 실행하여 매칭을 수행합니다.
3. 결과는 `LIFE_Auto_Table.xlsx` 파일에 저장됩니다.

## 요구 사항
- Python
- openpyxl
- pandas

## 설치
```bash
pip install openpyxl pandas
```

## 실행
```bash
runtennis
```
