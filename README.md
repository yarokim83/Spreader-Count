# ARMGC Spreader Count 자동 산정 및 DB 이력 관리 시스템 (v1.0.0)

ARMGC 38대 크레인(211~272호)에 장착되는 Spreader(Y01~Y44)의 교체 이력 및 **TLC(트위스트락횟수) 기반 누적 작업수**를 자동 산정하고, 로컬 SQLite DB(`spreader_history.db`)를 통해 **조회기간별 Spreader 사용 카운트**를 관리하는 데스크톱 프로그램입니다.

## 주요 기능
1. **대시보드 장착 현황 비교 및 교환일 자동 산정**
   - 두 시점의 대시보드 현황 사진(제미나이 OCR 판독 데이터)을 비교하여 Spreader가 교체된 호기를 자동 감지합니다.
   - **교환일 미상 시**: 산정 기간의 정중앙일(50:50 균등 배분)을 교환일로 자동 산정합니다.
   - **교환일 확인 시**: 수동 입력된 교환일을 기준으로 기간 내 장착 일수 비율(일할 계산)로 정밀 배분합니다.
2. **TLC (트위스트락횟수) 일괄 붙여넣기 연동**
   - 터미널 시스템/엑셀에서 복사한 `TC4211 ~ TC4272` 전체 표(`Eqp No | Work Time | MMBF | MTTR | TLC | Work Count`)를 붙여넣으면 자동으로 `TLC` 열을 추출하여 반영합니다.
3. **SQLite 로컬 DB(`spreader_history.db`) 이력 관리 및 기간별 사용량 조회**
   - 산정 회차별 데이터를 로컬 DB에 저장·불러오기·관리할 수 있습니다.
   - **조회기간(시작일 ~ 종료일)**을 설정하면 해당 기간 동안 각 Spreader 번호별 사용 카운트(TLC)와 교체 이력이 순위별로 집계됩니다.
4. **원본 관리 엑셀 양식 완벽 호환 내보내기**
   - `ARMGC Spreader Count List` 양식(좌우 2단 테이블 및 평균 수식) 그대로 엑셀 파일을 생성하며, 두 번째 시트에 호기별 교환일 산정 상세 근거를 함께 저장합니다.

## 파일 구성
- `spreader_count_app.py`: 메인 Python GUI 애플리케이션 소스코드
- `ARMGC_Spreader_Count_Manager.exe`: 파이썬 설치 없이 실행 가능한 Windows 단독 실행 파일
- `spreader_history.db`: 회차별 산정 이력 및 기간별 조회용 SQLite 데이터베이스
- `ARMGC_Spreader_Count_20260831.xlsx`: 기준 관리 엑셀 파일 (2026-08-31)
- `ARMGC_Spreader_Count_20261019.xlsx`: 산정 완료 관리 엑셀 파일 (2026-10-19, TLC 기준)
- `run_spreader_app.bat`: Python 스크립트 간편 실행 배치 파일

## 버전 이력 (Changelog)
- **v1.0.0 (2026-10-19)**
  - 초기 릴리즈: 2026-08-31 관리 엑셀 및 2026-09-14 ~ 2026-10-19 대시보드 OCR 판독 데이터 반영
  - TLC(트위스트락횟수) 기준 자동/수동 일할 배분 엔진 구현
  - SQLite DB(`spreader_history.db`) 저장 및 조회기간별 Spreader 사용 카운트 조회 탭 추가
  - Windows 단독 실행 배포 파일(`ARMGC_Spreader_Count_Manager.exe`) 빌드
