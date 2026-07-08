# sc2026 한글게임 정리 화면

`한깨_재미재미 한글게임 통합.xlsx`를 정리하기 위한 정적 HTML 화면입니다.

- `index.html`: 검색, 필터, 행 색상 변경, 셀 편집, CSV 다운로드, Supabase 수동 저장 화면
- `hangul-game-data.js`: 엑셀에서 추출한 저장 대상 데이터
- `supabase-config.js`: Supabase Project URL / anon key 설정
- `supabase-sc2026.sql`: Supabase `sc2026` 테이블 생성 SQL

엑셀에서 제외한 컬럼은 QA 검증, 노출/사용 여부, 중복 체크, 담당자, 개발완료일 계열입니다.
원본 셀 색상은 `회색`, `노랑`, `초록`, `파랑`, `투명or흰색`으로 분류해 저장합니다.

Supabase 저장을 사용하려면 `supabase-sc2026.sql`을 Supabase SQL Editor에서 먼저 실행해야 합니다.
셀 수정값과 행 색상 변경값은 화면에 먼저 반영되고, `수정사항 저장` 버튼을 눌렀을 때 `sc2026` 테이블에 저장됩니다.
