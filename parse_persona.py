"""
나무위키 Limbus Company 인격 MHT 파일을 파싱해서 인격별 JSON을 생성하는 스크립트.

사용법:
    python3 parse_persona.py <mht파일경로> <출력디렉토리>

동작 개요:
1. MHT(멀티파트 mime) 파일에서 본문 text/html 파트를 추출
2. BeautifulSoup으로 태그 제거 후 순수 텍스트 라인 리스트로 변환
3. "[ 인격이름 ] 수감자이름" 헤더 패턴으로 인격 단위 블록 분리
   (E.G.O 블록은 헤더는 같은 패턴이지만, 이 스크립트에서는 목차의
   "2. 인격" 절 범위만 대상으로 함 — sinner_index 파일 참고)
4. 각 블록 안에서 섹션 키워드(스테이터스/내성 정보/기본 정보/인격 동기화/
   흐트러짐 구간/스킬명/패시브/정신력 조건/대사)를 기준으로 필드 추출
5. 대사 텍스트는 원문을 저장하지 않고, "카테고리 존재 여부"만 기록

주의:
- 나무위키 문서는 편집자마다 서식이 조금씩 다를 수 있어 100% 자동화는
  어려움. 파싱 실패 필드는 null 또는 "_raw" 키에 원본 라인을 남겨서
  수동 보정이 쉽도록 함.
- 이 스크립트는 "이상_LCB 수감자" 블록으로 1차 검증됨. 다른 수감자
  문서에 적용 시 섹션 라벨이 다르면 (예: 방어 스킬 이름이 "가드"가
  아닌 경우 등) 실패할 수 있으므로 실행 후 결과를 반드시 검수할 것.
"""

import email
import json
import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup

SIN_TYPES = ["나태", "탐식", "색욕", "분노", "우울", "허영", "질투", "오만"]
ATTACK_TYPES = ["참격", "관통", "타격"]


def load_mht_text_lines(mht_path: str) -> list[str]:
    with open(mht_path, "rb") as f:
        msg = email.message_from_binary_file(f)

    best_html = None
    best_len = 0
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            payload = part.get_payload(decode=True)
            if payload and len(payload) > best_len:
                best_html = payload.decode("utf-8", errors="ignore")
                best_len = len(payload)

    if not best_html:
        raise ValueError("본문 HTML 파트를 찾지 못했습니다.")

    soup = BeautifulSoup(best_html, "lxml")
    for tag in soup(["script", "style"]):
        tag.decompose()
    text = soup.get_text("\n")
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    return lines


def find_persona_blocks(lines: list[str]) -> list[tuple[str, str, int, int]]:
    """[ 이름 ] 수감자 헤더로 블록 시작 인덱스를 찾아 (이름, 수감자, 시작, 끝) 리스트 반환."""
    header_re = re.compile(r"^\[\s*(.+?)\s*\]\s*(\S+)$")
    starts = []
    for i, line in enumerate(lines):
        m = header_re.match(line)
        if m:
            # 진짜 인격/E.G.O 헤더는 이후 5줄 이내에 "기본 이미지"가 나옴
            # (상단 배너의 "[ xxx ] 수감자" 목록은 이 조건을 만족하지 않음)
            lookahead = lines[i + 1:i + 15]
            if "동기화 이미지" in lookahead:
                starts.append((i, m.group(1), m.group(2)))

    blocks = []
    for idx, (start, name, sinner) in enumerate(starts):
        end = starts[idx + 1][0] if idx + 1 < len(starts) else len(lines)
        blocks.append((name, sinner, start, end))
    return blocks


def extract_between(lines: list[str], start_kw: str, end_kws: list[str], search_from: int, search_to: int):
    """start_kw 라인 다음부터 end_kws 중 하나가 나오기 전까지 라인 반환."""
    try:
        s = lines.index(start_kw, search_from, search_to) + 1
    except ValueError:
        return [], -1
    e = search_to
    for kw in (end_kws or []):
        try:
            cand = lines.index(kw, s, search_to)
            e = min(e, cand)
        except ValueError:
            pass
    return lines[s:e], s


def parse_sync_stages(block: list[str], start: int, end: int) -> list[dict]:
    stages = []
    i = start
    while i < end:
        if block[i] == "동기화 단계":
            i += 1
            if i < end and block[i] == "▶":
                i += 1
            cost_parts = []
            # 비용은 "스토리 개방" 또는 "필요 끈 N" (+"자아 파편 N")
            while i < end and (block[i] == "스토리 개방" or block[i].startswith("필요 끈") or block[i].startswith("자아 파편")):
                cost_parts.append(block[i])
                i += 1
            effects = []
            while i < end and block[i] not in ("동기화 단계", "흐트러짐 구간"):
                effects.append(block[i])
                i += 1
            stages.append({
                "stage": len(stages) + 1,
                "cost": ", ".join(cost_parts) if cost_parts else None,
                "effects": effects,
            })
        else:
            i += 1
    return stages


def find_skill_boundaries(lines: list[str], start: int, end: int) -> list[tuple[int, int, bool]]:
    """블록 내 "공격 유형"/"수비 유형" 라벨 위치를 기준으로
    각 스킬의 (이름 시작 인덱스, 유형라벨 인덱스, is_defense) 리스트를 반환."""
    boundaries = []
    for i in range(start, end):
        if lines[i] in ("공격 유형", "수비 유형"):
            is_defense = lines[i] == "수비 유형"
            j = i - 1
            while j > start and re.match(r"^-?\d+\(?[+-]?\d*\)?$", lines[j]):
                j -= 1
            boundaries.append((j, i, is_defense))
    return boundaries


def parse_skill_block(lines: list[str], start: int, end: int, is_defense: bool) -> dict:
    """스킬/방어스킬 하나의 블록(이름 포함, [start, end) 범위 내로 제한됨)을 파싱."""
    name = lines[start]
    i = start + 1

    # HP 표기 라인들 (숫자(부호숫자) 패턴), attack skill은 1개, defense는 2개인 경우 있음
    hp_vals = []
    while i < end and re.match(r"^-?\d+\(?[+-]?\d*\)?$", lines[i]):
        hp_vals.append(lines[i])
        i += 1

    skill = {
        "name": name,
        "hpDisplay": hp_vals,
    }

    field_map = {
        "공격 유형": "attackType",
        "수비 유형": "defenseType",
        "죄악 속성": "sinType",
        "스킬 수량": "coinCount",
        "스킬 위력": "skillPower",
        "코인 위력": "coinPower",
        "공격 가중치": "attackWeight",
        "가중치": "attackWeight",
    }

    while i < end and lines[i] != "[ 코인별 효과 ]":
        label = lines[i]
        if label in field_map:
            i += 1
            if i < end:
                skill[field_map[label]] = lines[i]
                i += 1
            continue
        i += 1

    # 코인별 효과: "[스킬 N] 효과" 는 더미 placeholder이므로 무시,
    # 그 외 조건("사용시","적중시","합 승리시","앞면 적중시" 등) + 다음 줄들을 effect로 묶음
    coin_effects = []
    if i < end and lines[i] == "[ 코인별 효과 ]":
        i += 1
        j = i
        placeholder_re = re.compile(r"^\[.+\] 효과$")
        condition_re = re.compile(r"^\[.+시\]$")
        while j < end:
            if placeholder_re.match(lines[j]):
                j += 1
                continue
            if condition_re.match(lines[j]):
                cond = lines[j]
                j += 1
                effect_lines = []
                while j < end and not placeholder_re.match(lines[j]) and not condition_re.match(lines[j]):
                    effect_lines.append(lines[j])
                    j += 1
                coin_effects.append({"condition": cond, "effect": " ".join(effect_lines)})
                continue
            j += 1
        i = j
    skill["coinEffects"] = coin_effects
    return skill, i


def parse_persona_block(name: str, sinner: str, lines: list[str], start: int, end: int) -> dict:
    persona = {
        "name": name,
        "sinner": sinner,
        "_parse_warnings": [],
    }

    # 등급
    rarity_m = None
    for i in range(start, min(start + 40, end)):
        m = re.search(r"\((\d)성\)", lines[i])
        if m:
            rarity_m = int(m.group(1))
            break
    persona["rarity"] = rarity_m

    # 스테이터스
    stat_lines, s_idx = extract_between(lines, "최대 레벨 · 동기화 기준", ["내성 정보"], start, end)
    if len(stat_lines) >= 3:
        persona["stats"] = {
            "maxLevel": stat_lines[0],
            "speed": stat_lines[1],
            "hp": stat_lines[2],
        }
    else:
        persona["_parse_warnings"].append("stats_incomplete")

    # 내성 정보
    resist_lines, _ = extract_between(lines, "내성 정보", ["기본 정보"], start, end)
    resistance = {}
    ri = 0
    while ri + 2 < len(resist_lines) and ri < 12:
        atk_type = resist_lines[ri]
        if atk_type in ATTACK_TYPES:
            level = resist_lines[ri + 1]
            mult = resist_lines[ri + 2] if ri + 2 < len(resist_lines) else None
            resistance[atk_type] = {"level": level, "multiplier": mult}
            ri += 3
        else:
            ri += 1
    persona["resistance"] = resistance

    # 기본 정보
    basic_lines, _ = extract_between(lines, "기본 정보", ["인격 동기화"], start, end)
    basic_map = {
        "수감자": "sinnerConfirm",
        "시즌": "season",
        "인격 등급": "rarityRaw",
        "출시 시기": "releaseDate",
        "티켓 인사말": "ticketGreeting",
        "획득 방법": "acquireMethod",
        "특성 키워드": "keywords",
    }
    bi = 0
    while bi < len(basic_lines):
        label = basic_lines[bi]
        if label in basic_map:
            bi += 1
            if bi < len(basic_lines):
                val = basic_lines[bi]
                if basic_map[label] == "keywords":
                    val = [k.strip() for k in val.split(",")]
                persona[basic_map[label]] = val
                bi += 1
            continue
        bi += 1

    # 인격 동기화
    sync_lines, sync_start = extract_between(lines, "인격 동기화", ["흐트러짐 구간"], start, end)
    if sync_start != -1:
        persona["syncUpgrades"] = parse_sync_stages(lines, sync_start, lines.index("흐트러짐 구간", sync_start, end))

    # 흐트러짐 구간
    stagger_lines, _ = extract_between(lines, "흐트러짐 구간", None, start, end)
    if stagger_lines:
        persona["staggerThreshold"] = stagger_lines[:3]

    # 스킬들: "흐트러짐 구간" 이후 3개 값 다음부터 "패시브" 전까지
    try:
        stagger_idx = lines.index("흐트러짐 구간", start, end)
        skills_start = stagger_idx + 1 + 3  # 흐트러짐 구간 라벨 + 3개 퍼센트
    except ValueError:
        skills_start = None

    try:
        passive_idx = lines.index("패시브", start, end)
    except ValueError:
        passive_idx = end

    skills = []
    defense_skill = None
    if skills_start is not None:
        boundaries = find_skill_boundaries(lines, skills_start, passive_idx)
        for k, (name_idx, _type_idx, is_defense) in enumerate(boundaries):
            block_end = boundaries[k + 1][0] if k + 1 < len(boundaries) else passive_idx
            skill, _ = parse_skill_block(lines, name_idx, block_end, is_defense)
            if is_defense:
                defense_skill = skill
            else:
                skills.append(skill)
    persona["skills"] = skills
    persona["defenseSkill"] = defense_skill

    # 패시브
    passives = []
    panic_idx = end
    if passive_idx < end:
        i = passive_idx + 1
        support_idx = None
        try:
            support_idx = lines.index("서포트 패시브", passive_idx, end)
        except ValueError:
            pass
        panic_idx = None
        try:
            panic_idx = lines.index("패닉 유형", passive_idx, end)
        except ValueError:
            panic_idx = end

        def parse_one_passive(pname_idx, pend, ptype):
            pname = lines[pname_idx]
            j = pname_idx + 1
            resonance = lines[j] if j < pend else None
            j += 1
            # 0, 0 더미 스킵
            while j < pend and re.match(r"^-?\d+$", lines[j]):
                j += 1
            desc_lines = []
            while j < pend and lines[j] not in ("서포트 패시브",):
                desc_lines.append(lines[j])
                j += 1
            return {"name": pname, "type": ptype, "resonance": resonance, "description": " ".join(desc_lines)}, j

        if support_idx is not None:
            p1, _ = parse_one_passive(i, support_idx, "전투 패시브")
            p2, _ = parse_one_passive(support_idx + 1, panic_idx, "서포트 패시브")
            passives = [p1, p2]
        else:
            p1, _ = parse_one_passive(i, panic_idx, "전투 패시브")
            passives = [p1]
    persona["passives"] = passives

    # 정신력(패닉) 조건
    if panic_idx < end:
        try:
            panic_type = lines[panic_idx + 1]
            panic_effect = lines[panic_idx + 2]
            inc_idx = lines.index("정신력 증가 조건", panic_idx, end)
            dec_idx = lines.index("정신력 감소 조건", inc_idx, end)
            dialogue_idx = None
            try:
                # "대사" 섹션과 "OOO의 이야기" 섹션 순서가 인격마다 다름.
                # 둘 다 제목 줄 바로 다음에 "[ 펼치기 · 접기 ]" 가 오므로,
                # dec_idx 이후 첫 "[ 펼치기 · 접기 ]" 직전 줄(=섹션 제목)에서 자른다.
                fold_idx = lines.index("[ 펼치기 · 접기 ]", dec_idx, end)
                dialogue_idx = fold_idx - 1
            except ValueError:
                dialogue_idx = end
            persona["sanity"] = {
                "panicType": panic_type,
                "panicEffect": panic_effect,
                "increaseConditions": lines[inc_idx + 1:dec_idx],
                "decreaseConditions": lines[dec_idx + 1:dialogue_idx],
            }
        except ValueError:
            persona["_parse_warnings"].append("sanity_parse_failed")

    return persona


def sanitize_filename(name: str) -> str:
    """OS별 파일명 금지 문자(콜론, 슬래시 등)를 안전한 문자로 치환."""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def main():
    if len(sys.argv) >= 3:
        mht_path = sys.argv[1]
        out_dir = Path(sys.argv[2])
        name_filter = sys.argv[3] if len(sys.argv) > 3 else None
    else:
        print("=== 나무위키 인격 파서 ===")
        mht_path = input("MHT 파일 경로를 입력하세요: ").strip().strip('"')
        out_dir_input = input("출력 디렉토리 (엔터 시 ./output): ").strip()
        out_dir = Path(out_dir_input) if out_dir_input else Path("./output")
        name_filter_input = input("특정 인격 이름만 파싱하려면 입력 (전체는 엔터): ").strip()
        name_filter = name_filter_input if name_filter_input else None

    if not Path(mht_path).exists():
        print(f"파일을 찾을 수 없습니다: {mht_path}")
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)

    lines = load_mht_text_lines(mht_path)
    blocks = find_persona_blocks(lines)

    results = []
    for name, sinner, start, end in blocks:
        if name_filter and name != name_filter:
            continue
        persona = parse_persona_block(name, sinner, lines, start, end)
        out_path = out_dir / sanitize_filename(f"{sinner}_{name}.json".replace(" ", "_"))
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(persona, f, ensure_ascii=False, indent=2)
        results.append(str(out_path))

    print(f"{len(results)}개 인격 블록 파싱 완료")
    for r in results:
        print(" -", r)


if __name__ == "__main__":
    main()
