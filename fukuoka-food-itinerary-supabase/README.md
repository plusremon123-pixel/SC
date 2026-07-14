# 후쿠오카 식당 선택 Supabase 버전

기존 `개정 일정 html만들기` 프로젝트의 Supabase Storage 정보를 사용합니다.

- Supabase URL: `https://vmebzlinboxmgcrrorwv.supabase.co`
- Bucket: `schedule-data`
- 저장 파일: `fukuoka-food-itinerary-state.json`

기본 식당/디저트/매장 정보는 `index.html` 안에 들어 있고, 사용자가 바꾼 선택값과 추천 다시하기 결과는 Supabase Storage JSON으로 저장됩니다.

브라우저가 오프라인이거나 Supabase 저장이 실패하면 같은 브라우저의 `localStorage` 값을 우선 유지합니다.
