# Known Issues & Limitations

누적된 구조적 한계 및 미해결 이슈를 추적합니다.  
각 항목은 고유 ID로 조회 가능합니다.

---

## F (Structural Flaw)

### F1 — 들여쓰기 계층 미반영 (Presentation Linkbase 미제공)

**현상**  
재무제표 출력 시 계정 항목의 들여쓰기가 실제 위계를 반영하지 못함.  
예: 삼성전자 포괄손익계산서에서 "후속적으로 재분류되지 않는 포괄손익"이  
"연결기타포괄손익"의 자식 항목임을 출력물에서 알 수 없음.

**원인**  
DART `fnlttSinglAcntAll` API는 XBRL Instance Document의 값(숫자)만 반환하며,  
계층 구조를 정의하는 **Presentation Linkbase**를 제공하지 않음.  
→ 부모-자식 관계 및 들여쓰기 깊이 정보 소실 (Structural flaw)

**현재 대응 (B안)**  
단순 2단계 indent로 근사:
- Level 3/1 (합계·헤더): indent 0
- Level 2 (섹션 소계): indent 1
- Level 0 (세부항목): indent 2

**근본 해결 (A안, 미적용)**  
XBRL 원문 파일 다운로드 후 Presentation Linkbase 직접 파싱.  
단, 회사별·연도별 구조 차이로 과적합 위험 존재.

**관련 재무제표**: IS, CIS, CF, SCE (BS는 단순 구조라 영향 적음)

---
