"""recipes.py — 레시피 지침문 로더·매칭(PRD 16-1 ③, 2026-10-02 T6 신설).

"자주 하는 업무는 코드가 아니라 레시피 지침문으로 힌트를 준다"의 구현.
recipes/ 폴더의 .md 파일 하나가 레시피 하나이며, 형식은 사람이 그냥 읽고
고칠 수 있는 마크다운이다(사업 카드와 같은 원칙 — 사용자가 직접 수정하는
파일). 매칭은 LLM이 아니라 결정적 키워드 포함 검사다 — 소형 모델(Qwen)에
기대지 않는 안전망이라는 16절 원칙.

레시피 파일 형식(## 머리글로 구획):
    # 레시피: <이름>
    ## 트리거
    쉼표로, 구분된, 키워드들
    ## 단계
    1. 첫 단계
    2. 둘째 단계
    ## 지침
    - 지침 한 줄
    - 지침 또 한 줄

루프와의 연결(agent_loop.build_graph의 recipe_matcher 인자):
- 문맥 수집 노드가 match_recipe(요청문)를 불러 매칭되면 context["recipe"]에
  넣는다 → OllamaPlanner.plan의 프롬프트에 문맥으로 함께 들어가 계획 힌트가
  된다.
- 계획 노드는 LLM 계획이 실패하거나 비면 레시피의 고정 단계를 그대로 쓴다
  ("9B가 흔들리면 레시피 기반 고정 목록으로 대체", PRD 16-3)."""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
RECIPES_DIR = os.path.join(HERE, "recipes")

_NUMBERED = re.compile(r"^\s*\d+[.)]\s*(.+)$")
_BULLET = re.compile(r"^\s*[-•]\s*(.+)$")


def parse_recipe(text: str, name_fallback: str = "") -> dict:
    """레시피 마크다운 한 편을 {"name", "triggers", "steps", "guidance"}로
    파싱한다. 구획이 없거나 비어 있으면 빈 리스트로 둔다(파일 하나가 깨져도
    전체 로딩이 죽지 않게 — load_recipes에서 트리거 없는 레시피는 버린다)."""
    name = name_fallback
    m = re.search(r"^#\s*레시피\s*[:：]\s*(.+)$", text, re.MULTILINE)
    if m:
        name = m.group(1).strip()
    sections = {}
    current = None
    for line in text.splitlines():
        h = re.match(r"^##\s*(\S+)", line)
        if h:
            current = h.group(1)
            sections[current] = []
        elif current is not None:
            sections[current].append(line)

    triggers = []
    for line in sections.get("트리거", []):
        triggers += [t.strip() for t in line.split(",") if t.strip()]
    steps = [m.group(1).strip() for ln in sections.get("단계", []) if (m := _NUMBERED.match(ln))]
    guidance = [m.group(1).strip() for ln in sections.get("지침", []) if (m := _BULLET.match(ln))]
    return {"name": name, "triggers": triggers, "steps": steps, "guidance": guidance}


def load_recipes(recipes_dir: str = RECIPES_DIR) -> list[dict]:
    """recipes/의 .md 전부를 파싱해 돌려준다(파일명 순 — 매칭 동점 시 순서가
    결정적이어야 해서). 폴더가 없으면 빈 리스트. 트리거가 없는 레시피는
    매칭될 수 없으므로 제외한다."""
    if not os.path.isdir(recipes_dir):
        return []
    recipes = []
    for fname in sorted(os.listdir(recipes_dir)):
        if not fname.lower().endswith(".md"):
            continue
        path = os.path.join(recipes_dir, fname)
        try:
            with open(path, encoding="utf-8") as f:
                recipe = parse_recipe(f.read(), name_fallback=os.path.splitext(fname)[0])
        except Exception:  # noqa: BLE001 — 한 파일이 깨져도 나머지는 산다
            continue
        if recipe["triggers"]:
            recipes.append(recipe)
    return recipes


def match_recipe(user_message: str, recipes: list[dict] = None) -> dict | None:
    """요청문에 포함된 트리거 키워드가 가장 많은 레시피를 돌려준다(결정적 —
    LLM 안 씀). 하나도 안 걸리면 None. 동점이면 파일명 순으로 앞선 것."""
    if recipes is None:
        recipes = load_recipes()
    best, best_hits = None, 0
    for recipe in recipes:
        hits = sum(1 for t in recipe["triggers"] if t in user_message)
        if hits > best_hits:
            best, best_hits = recipe, hits
    return best


# ---------------------------------------------------------------- self-tests (서버·한글 불필요)

_SAMPLE = """# 레시피: 틀 유지 초안
설명 줄은 무시된다.
## 트리거
틀 유지, 틀을 유지, 참고해서 초안
## 단계
1. 구간 목록을 읽는다
2. 첨부를 읽는다
3) 구간마다 고쳐 쓴다
## 지침
- 제목은 그대로 둔다
- 수치를 지어내지 않는다
"""


def _selftest_parse_recipe():
    r = parse_recipe(_SAMPLE, name_fallback="파일명")
    assert r["name"] == "틀 유지 초안", r
    assert r["triggers"] == ["틀 유지", "틀을 유지", "참고해서 초안"], r
    assert r["steps"] == ["구간 목록을 읽는다", "첨부를 읽는다", "구간마다 고쳐 쓴다"], r
    assert r["guidance"] == ["제목은 그대로 둔다", "수치를 지어내지 않는다"], r
    empty = parse_recipe("# 레시피: 깨진 파일\n본문뿐", name_fallback="x")
    assert empty["triggers"] == [] and empty["steps"] == [], empty
    print("parse_recipe 통과")


def _selftest_match_recipe():
    recipes = [parse_recipe(_SAMPLE),
               {"name": "회의 반영", "triggers": ["회의결과", "반영"], "steps": ["s"], "guidance": []}]
    r = match_recipe("작년 계획안 틀을 유지해서 만들어줘", recipes)
    assert r and r["name"] == "틀 유지 초안", r
    r = match_recipe("회의결과를 반영해줘", recipes)  # 트리거 2개 적중 > 0개
    assert r and r["name"] == "회의 반영", r
    assert match_recipe("숫자 검증해줘", recipes) is None
    print("match_recipe 통과")


def _selftest_load_recipes_real_folder():
    """실제 recipes/ 폴더가 있으면 전부 트리거·단계를 갖춰야 한다(깨진 레시피
    조기 발견). 폴더가 없으면 빈 리스트로 조용히 동작."""
    recipes = load_recipes()
    for r in recipes:
        assert r["triggers"] and r["steps"], r["name"]
    print(f"load_recipes 통과({len(recipes)}개 로딩)")


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    _selftest_parse_recipe()
    _selftest_match_recipe()
    _selftest_load_recipes_real_folder()
