import streamlit as st
import pandas as pd
import numpy as np
import re
import io

st.set_page_config(page_title="명부정리 & 표본추출", page_icon="📋", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.title-box {
    background: linear-gradient(135deg, #1a3c6e 0%, #2563b0 100%);
    color: white; padding: 2rem 2.5rem; border-radius: 12px; margin-bottom: 2rem;
}
.title-box h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
.title-box p  { margin: 0.4rem 0 0; opacity: 0.85; font-size: 0.95rem; }
.info-box {
    background: #f0f6ff; border-left: 4px solid #2563b0;
    padding: 0.9rem 1.2rem; border-radius: 6px;
    font-size: 0.88rem; color: #334155; margin-bottom: 1rem;
}
.warn-box {
    background: #fffbeb; border-left: 4px solid #f59e0b;
    padding: 0.9rem 1.2rem; border-radius: 6px;
    font-size: 0.88rem; color: #334155; margin-bottom: 1rem;
}
.stat-card {
    background: white; border: 1px solid #e2e8f0; border-radius: 10px;
    padding: 1.2rem; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
.stat-num   { font-size: 1.8rem; font-weight: 700; }
.stat-label { font-size: 0.78rem; color: #64748b; margin-top: 0.2rem; }
.c-blue  { color: #1e40af; }
.c-green { color: #16a34a; }
.c-red   { color: #dc2626; }
.c-gray  { color: #475569; }
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# 비밀번호 인증
# ════════════════════════════════════════════════════════════
PASSWORD = "1974"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="title-box">
        <h1>📋 명부정리 & 표본추출 시스템</h1>
        <p>접근하려면 비밀번호를 입력하세요</p>
    </div>
    """, unsafe_allow_html=True)
    pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
    if st.button("확인", type="primary"):
        if pw == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    st.stop()

# ════════════════════════════════════════════════════════════
# 인증 통과 후 메인 화면
# ════════════════════════════════════════════════════════════
st.markdown("""
<div class="title-box">
    <h1>📋 명부정리 & 표본추출 시스템</h1>
    <p>명부정리 완료 후 파일을 확인하고, 표본추출 탭에서 추출을 진행하세요</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📁 1단계: 명부정리", "🎯 2단계: 표본추출"])

# ════════════════════════════════════════════════════════════
# 공통 유틸
# ════════════════════════════════════════════════════════════
def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def to_excel_multi(sheets: dict):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
    return buf.getvalue()

def smart_read(f):
    """빈 행을 건너뛰고 실제 헤더 행 자동 탐지"""
    df_peek = pd.read_excel(f, header=None, nrows=10)
    header_row = 0
    for i, row in df_peek.iterrows():
        if row.notna().sum() > 3:
            header_row = i
            break
    f.seek(0)
    df = pd.read_excel(f, header=header_row)
    df.columns = df.columns.astype(str).str.strip()
    return df

def find_col(df, keywords):
    for kw in keywords:
        matches = [c for c in df.columns if kw in c]
        if matches:
            return matches[0]
    return None

# ════════════════════════════════════════════════════════════
# 전화번호 정리
# ════════════════════════════════════════════════════════════
def format_phone(raw):
    if pd.isna(raw):
        return ""
    s = str(raw).strip()
    if "*" in s or "＊" in s:
        return ""
    digits = re.sub(r"\D", "", s)
    if not digits or len(digits) < 7:
        return ""
    if digits.startswith("02"):
        local = digits[2:]
        if len(local) == 7:   return f"02-{local[:3]}-{local[3:]}"
        elif len(local) == 8: return f"02-{local[:4]}-{local[4:]}"
        else:                 return f"02-{local}"
    elif len(digits) == 10:   return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    elif len(digits) == 11:   return f"{digits[:3]}-{digits[3:7]}-{digits[7:]}"
    else:                     return digits

# ════════════════════════════════════════════════════════════
# 주소 조합
# ════════════════════════════════════════════════════════════
def build_address(row):
    def s(v):
        if pd.isna(v): return ""
        v = str(v).strip()
        if v in ("nan", "None", "NaN", ""): return ""
        try:
            f = float(v)
            if f == int(f): return str(int(f))
            return v
        except:
            return v

    haengjeong  = s(row.get("행정구역", ""))
    road        = s(row.get("도로명", ""))
    underground = s(row.get("도로명지하", ""))
    main_no     = s(row.get("도로명본번지", ""))
    sub_no      = s(row.get("도로명부번지", ""))
    bldg        = s(row.get("도로명건물명(빌딩시장상가)", ""))
    dong        = s(row.get("도로명건물동", ""))
    floor       = s(row.get("도로명건물층", ""))
    ho          = s(row.get("도로명건물호", ""))

    inaccurate = False

    def drop_last_word(text):
        parts = text.split()
        return " ".join(parts[:-1]) if len(parts) >= 3 else text

    if road and main_no:
        prefix = drop_last_word(haengjeong) if haengjeong else ""
        underground_str = "지하" if underground == "1" else ""
        if sub_no and sub_no != "0":
            road_no = f"{underground_str}{main_no}-{sub_no}"
        else:
            road_no = f"{underground_str}{main_no}"
        parts = [p for p in [prefix, road, road_no] if p]
        addr  = " ".join(parts)
        detail = []
        if bldg:  detail.append(bldg)
        if dong:  detail.append(dong + "동")
        if floor: detail.append(floor + "층")
        if ho:    detail.append(ho + "호")
        if detail: addr += " " + " ".join(detail)
        if not prefix or not road or not main_no: inaccurate = True
    else:
        addr = haengjeong
        detail = []
        if bldg:  detail.append(bldg)
        if dong:  detail.append(dong + "동")
        if floor: detail.append(floor + "층")
        if ho:    detail.append(ho + "호")
        if detail: addr += " " + " ".join(detail)
        if not haengjeong: inaccurate = True

    if not addr.strip():
        return "", True
    return addr.strip(), inaccurate

# ════════════════════════════════════════════════════════════
# 할당표본 파일 자동 파싱
# ════════════════════════════════════════════════════════════
def parse_allocation(f):
    """건설업/용역업/제조업이 혼재된 할당표본 파일을 자동 파싱"""
    df = pd.read_excel(f, header=None)
    results = []
    i = 0

    while i < len(df):
        cell = str(df.iloc[i, 0]).strip()

        if any(k in cell for k in ["건설업", "용역업", "제조업"]):
            if "건설업" in cell:   industry = "건설업"
            elif "용역업" in cell: industry = "용역업"
            else:                  industry = "제조업"

            nat_start = pub_start = None
            mid_row   = None

            # 전국/공시 위치 + 중분류 헤더 찾기
            for j in range(i + 1, min(i + 6, len(df))):
                row = df.iloc[j]
                for col_idx, val in enumerate(row):
                    if str(val).strip() == "전국" and nat_start is None:
                        nat_start = col_idx
                    if str(val).strip() == "공시" and pub_start is None:
                        pub_start = col_idx
                if str(row.iloc[0]).strip() == "중분류":
                    mid_row = j

            if mid_row is None or nat_start is None or pub_start is None:
                i += 1
                continue

            # 규모층 번호 탐지 (중분류 행 또는 바로 다음 행)
            size_nat, size_pub = [], []
            data_start = mid_row + 1

            for size_row_idx in [mid_row, mid_row + 1]:
                size_row = df.iloc[size_row_idx]
                sn = [(col_idx, str(v).strip().replace("*", ""))
                      for col_idx, v in enumerate(size_row)
                      if nat_start <= col_idx < pub_start
                      and str(v).strip().replace("*", "").isdigit()]
                sp = [(col_idx, str(v).strip().replace("*", ""))
                      for col_idx, v in enumerate(size_row)
                      if col_idx >= pub_start
                      and str(v).strip().replace("*", "").isdigit()]
                if sn or sp:
                    size_nat   = sn
                    size_pub   = sp
                    data_start = size_row_idx + 1
                    break

            # 데이터 행 읽기
            for k in range(data_start, len(df)):
                row     = df.iloc[k]
                mid_val = str(row.iloc[0]).strip()
                if pd.isna(row.iloc[0]) or mid_val in ("", "nan", "합계", "총합계", "NaN"):
                    if mid_val in ("합계", "총합계"): break
                    continue
                if any(kw in mid_val for kw in ["건설업", "용역업", "제조업"]): break

                for col_idx, size in size_nat:
                    try:
                        n = int(float(row.iloc[col_idx]))
                        if n > 0:
                            results.append({"업종": industry, "구분": "전국",
                                            "중분류": mid_val, "규모층": size, "할당표본": n})
                    except: pass

                for col_idx, size in size_pub:
                    try:
                        n = int(float(row.iloc[col_idx]))
                        if n > 0:
                            results.append({"업종": industry, "구분": "공시",
                                            "중분류": mid_val, "규모층": size, "할당표본": n})
                    except: pass
        i += 1

    return pd.DataFrame(results)

# ════════════════════════════════════════════════════════════
# TAB 1: 명부정리
# ════════════════════════════════════════════════════════════
with tab1:
    st.subheader("모집단 파일 업로드")
    st.markdown('<div class="info-box">정리할 모집단 xlsx 파일을 업로드하세요. (전국사업체조사 또는 공시대상)</div>', unsafe_allow_html=True)
    raw_file = st.file_uploader("모집단 파일", type=["xlsx"], key="raw")

    if raw_file:
        try:
            df_raw = smart_read(raw_file)
            st.success(f"✅ 파일 로드 완료 — {len(df_raw):,}개 행, {len(df_raw.columns)}개 열")
        except Exception as e:
            st.error(f"파일 읽기 오류: {e}")
            df_raw = None
    else:
        df_raw = None

    if df_raw is not None and st.button("🔧 명부정리 시작", type="primary"):

        df = df_raw.copy()
        df.columns = df.columns.astype(str).str.strip()
        df["_삭제사유"] = ""

        # ── 1. 중복 사업체 제거 ──────────────────────────────────────
        st.info("① 중복 사업체 제거 중...")

        INDUSTRY_GROUP = {
            "C": "제조업", "G": "제조업",
            "F": "건설업",
            "H": "용역업", "J": "용역업", "L": "용역업",
            "M": "용역업", "N": "용역업"
        }

        def get_industry_group(code):
            return INDUSTRY_GROUP.get(str(code).strip().upper(), str(code).strip().upper())

        biz_col      = find_col(df, ["사업자등록번호"])
        industry_col = find_col(df, ["산업대분류코드"])

        if biz_col:
            df[biz_col] = df[biz_col].astype(str).str.strip()
            st.write(f"   → 사업자등록번호 열: '{biz_col}'")

            if industry_col:
                df["_업종분야"] = df[industry_col].apply(get_industry_group)
            else:
                df["_업종분야"] = "기타"
                st.warning("'산업대분류코드' 열을 찾을 수 없어 업종 구분 없이 처리합니다.")

            def pick_survivor(group):
                code2 = group[group["사업체구분코드"].astype(str) == "2"]
                if len(code2) == 1: return code2.index[0]
                candidates = code2 if len(code2) > 0 else group
                candidates = candidates.copy()
                candidates["_종사자"] = pd.to_numeric(candidates["종사자규모"], errors="coerce").fillna(0)
                max_emp    = candidates["_종사자"].max()
                candidates = candidates[candidates["_종사자"] == max_emp]
                if len(candidates) == 1: return candidates.index[0]
                candidates["_매출액"] = pd.to_numeric(candidates["매출액규모"], errors="coerce").fillna(0)
                return candidates["_매출액"].idxmax()

            dup_mask = df.duplicated(subset=[biz_col, "_업종분야"], keep=False)
            dup_df   = df[dup_mask].copy()
            survivors = set()

            for (biz, field), group in dup_df.groupby([biz_col, "_업종분야"]):
                if industry_col:
                    code_counts    = group[industry_col].value_counts()
                    dominant_code  = code_counts.index[0]
                    dominant_group = group[group[industry_col] == dominant_code]
                    survivors.add(pick_survivor(dominant_group))
                else:
                    survivors.add(pick_survivor(group))

            dup_remove_idx = dup_df.index.difference(list(survivors))
            df.loc[dup_remove_idx, "_삭제사유"] = "중복(사업자등록번호+업종분야)"
            df = df.drop(columns=["_업종분야"])
            st.write(f"   → 중복 제거 대상: {len(dup_remove_idx):,}개")
        else:
            st.warning("'사업자등록번호' 열을 찾을 수 없습니다.")

        # ── 2. 전화번호 정리 ─────────────────────────────────────────
        st.info("② 전화번호 정리 중...")
        phone_col = find_col(df, ["전화번호"])
        if phone_col:
            df[phone_col] = df[phone_col].apply(format_phone)
            st.write(f"   → 전화번호 열: '{phone_col}' 정리 완료")
        else:
            st.warning("'전화번호' 열을 찾을 수 없습니다.")

        # ── 3. 주소 조합 ──────────────────────────────────────────────
        st.info("③ 주소 조합 중...")
        addr_results   = df.apply(build_address, axis=1)
        df["통합주소"]  = addr_results.apply(lambda x: x[0] + " (부정확)" if x[1] else x[0])
        inaccurate_cnt = addr_results.apply(lambda x: x[1]).sum()

        # 통합주소를 도로명건물호 바로 다음에 삽입
        addr_cols = ["시도","시군구","읍면동","행정구역","도로명읍면동","도로명",
                     "도로명지하","도로명본번지","도로명부번지",
                     "도로명건물명(빌딩시장상가)","도로명건물동","도로명건물층","도로명건물호"]
        existing_addr = [c for c in addr_cols if c in df.columns]
        if existing_addr:
            last_col   = existing_addr[-1]
            insert_pos = list(df.columns).index(last_col) + 1
            cols       = list(df.columns)
            cols.remove("통합주소")
            cols.insert(insert_pos, "통합주소")
            df = df[cols]
        st.write(f"   → 주소 조합 완료 (부정확 표시: {inaccurate_cnt:,}개)")

        # ── 결과 분리 ─────────────────────────────────────────────────
        delete_mask  = df["_삭제사유"] != ""
        df_clean     = df[~delete_mask].drop(columns=["_삭제사유"]).reset_index(drop=True)
        df_original  = df.copy()
        df_original["삭제여부"] = df_original["_삭제사유"].apply(lambda x: "삭제" if x else "")
        df_original  = df_original.drop(columns=["_삭제사유"]).reset_index(drop=True)

        # ── 요약 ──────────────────────────────────────────────────────
        st.markdown("### 📊 명부정리 결과")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-blue">{len(df_raw):,}</div><div class="stat-label">원본 행 수</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-green">{len(df_clean):,}</div><div class="stat-label">정리 후 (깨끗한 명부)</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-red">{delete_mask.sum():,}</div><div class="stat-label">삭제 대상</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        preview_cols = ["PID","사업체명","사업자등록번호","통합주소","전화번호","사업체구분코드"]
        preview_cols = [c for c in preview_cols if c in df_clean.columns]
        st.markdown("**미리보기 (상위 5건 — 깨끗한 명부)**")
        st.dataframe(df_clean[preview_cols].head(5), use_container_width=True)

        st.markdown("### 📥 다운로드")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label="✅ 깨끗한 명부 다운로드",
                data=to_excel_bytes(df_clean),
                file_name="명부정리_깨끗한명부.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                label="📋 전체 원본 명부 다운로드",
                data=to_excel_bytes(df_original),
                file_name="명부정리_전체원본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.markdown('<div class="warn-box">💡 깨끗한 명부를 확인한 후 <b>2단계 표본추출</b> 탭으로 이동하세요.</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# TAB 2: 표본추출
# ════════════════════════════════════════════════════════════
with tab2:
    st.subheader("정리된 모집단 파일 업로드")
    st.markdown("""
    <div class="info-box">
    • <b>전국사업체조사</b>: 1단계 명부정리 후 다운받은 <b>깨끗한 명부</b>를 업로드하세요.<br>
    • <b>공시대상</b>: 별도 정리 없이 <b>원본 파일을 바로</b> 업로드하세요.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        file_nat = st.file_uploader("전국사업체조사 깨끗한 명부", type=["xlsx"], key="s_nat")
    with col2:
        file_pub = st.file_uploader("공시대상 원본 파일", type=["xlsx"], key="s_pub")

    st.subheader("할당표본 파일 업로드")
    st.markdown('<div class="info-box">건설업/용역업/제조업 표본할당결과가 담긴 xlsx 파일을 업로드하세요. (기존 양식 그대로 업로드 가능)</div>', unsafe_allow_html=True)
    file_alloc = st.file_uploader("할당표본 파일", type=["xlsx"], key="s_alloc")

    ready2 = file_nat and file_pub and file_alloc

    if st.button("🎯 표본추출 시작", type="primary", disabled=not ready2):

        # ── 파일 로드 ──
        try:
            df_nat = smart_read(file_nat)
            df_pub = smart_read(file_pub)
            st.success(f"✅ 전국: {len(df_nat):,}개 / 공시: {len(df_pub):,}개")
        except Exception as e:
            st.error(f"모집단 파일 로드 오류: {e}"); st.stop()

        # ── 할당표본 파싱 ──
        try:
            alloc_long = parse_allocation(file_alloc)
            if alloc_long.empty:
                st.error("할당표본 파일 파싱 실패: 건설업/용역업/제조업 헤더를 찾을 수 없습니다.")
                st.stop()
            st.success(f"✅ 할당표본 파싱 완료 — {len(alloc_long)}개 층")
        except Exception as e:
            st.error(f"할당표본 파싱 오류: {e}"); st.stop()

        # ── 전처리 ──
        for df_ in [df_nat, df_pub]:
            df_["산업중분류코드"] = df_["산업중분류코드"].astype(str).str.strip() if "산업중분류코드" in df_.columns else ""
            df_["매출액규모"]     = df_["매출액규모"].astype(str).str.strip()     if "매출액규모"     in df_.columns else ""
            df_["산업분류코드"]   = df_["산업분류코드"].astype(str).str.strip()   if "산업분류코드"   in df_.columns else ""
            df_["종사자규모"]     = pd.to_numeric(df_["종사자규모"], errors="coerce").fillna(0) if "종사자규모" in df_.columns else 0
            df_["도로명주소"]     = df_.get("통합주소", df_.get("도로명주소", pd.Series([""] * len(df_)))).astype(str)
            df_["사업체구분코드"] = df_["사업체구분코드"].astype(str).str.strip() if "사업체구분코드" in df_.columns else ""

        pub_pids = set(df_pub["PID"].astype(str).tolist())

        # ── 계통추출 ──
        def systematic_sample(group_df, n):
            N = len(group_df)
            if N == 0 or n == 0: return []
            if n >= N: return list(group_df.index)
            interval  = N / n
            start     = np.random.uniform(1, interval + 1)
            positions = [min(int(start + interval * k) - 1, N - 1) for k in range(n)]
            return list(group_df.iloc[positions].index)

        def replace_dup(sample_indices, group_df, dup_pids):
            group_indices = list(group_df.index)
            sample_set    = set(sample_indices)
            result        = []
            for idx in sample_indices:
                pid = str(group_df.loc[idx, "PID"])
                if pid in dup_pids:
                    pos, replaced = group_indices.index(idx), False
                    for delta in [1, -1, 2, -2, 3, -3]:
                        np_ = pos + delta
                        if 0 <= np_ < len(group_indices):
                            ni  = group_indices[np_]
                            np2 = str(group_df.loc[ni, "PID"])
                            if ni not in sample_set and np2 not in dup_pids:
                                result.append(ni); sample_set.add(ni); replaced = True; break
                    if not replaced: result.append(idx)
                else:
                    result.append(idx); sample_set.add(idx)
            return result

        def extract_layer(layer_df, n, dup_pids):
            """코드 1,2 우선 추출 → 부족하면 3에서 보충"""
            layer_12 = layer_df[layer_df["사업체구분코드"].isin(["1", "2"])].copy()
            layer_3  = layer_df[layer_df["사업체구분코드"] == "3"].copy()
            N12 = len(layer_12)
            if N12 >= n:
                idx = systematic_sample(layer_12.reset_index(drop=False).set_index("index"), n)
                idx = replace_dup(idx, layer_12.reset_index(drop=False).set_index("index"), dup_pids)
                return idx, n, 0
            else:
                idx_12 = list(layer_12.index)
                need   = n - N12
                idx_3  = systematic_sample(layer_3.reset_index(drop=False).set_index("index"), need)
                idx_3  = replace_dup(idx_3, layer_3.reset_index(drop=False).set_index("index"), dup_pids)
                return idx_12 + idx_3, N12, len(idx_3)

        # ── 전국 추출 ──
        st.info("전국사업체조사 표본추출 중...")
        alloc_nat   = alloc_long[alloc_long["구분"] == "전국"]
        nat_samples, nat_log = [], []

        for _, row in alloc_nat.iterrows():
            mid, size, n = str(row["중분류"]), str(row["규모층"]), int(row["할당표본"])
            layer = df_nat[
                (df_nat["산업중분류코드"] == mid) &
                (df_nat["매출액규모"]    == size)
            ].copy()
            layer = layer.sort_values(["산업분류코드", "종사자규모", "도로명주소"])
            N     = len(layer)
            nat_log.append({"업종": row["업종"], "중분류": mid, "규모층": size, "모집단": N, "할당표본": n})
            if N == 0: continue
            idx, from_12, from_3 = extract_layer(layer, n, pub_pids)
            nat_samples.extend(idx)
            nat_log[-1]["코드1·2에서"] = from_12
            nat_log[-1]["코드3에서"]   = from_3

        # ── 공시 추출 ──
        st.info("공시대상 표본추출 중...")
        alloc_pub   = alloc_long[alloc_long["구분"] == "공시"]
        pub_samples, pub_log = [], []

        for _, row in alloc_pub.iterrows():
            mid, size, n = str(row["중분류"]), str(row["규모층"]), int(row["할당표본"])
            layer = df_pub[
                (df_pub["산업중분류코드"] == mid) &
                (df_pub["매출액규모"]    == size)
            ].copy()
            layer = layer.sort_values(["산업분류코드", "종사자규모", "도로명주소"])
            N     = len(layer)
            pub_log.append({"업종": row["업종"], "중분류": mid, "규모층": size, "모집단": N, "할당표본": n})
            if N == 0: continue
            idx, from_12, from_3 = extract_layer(layer, n, set())
            pub_samples.extend(idx)
            pub_log[-1]["코드1·2에서"] = from_12
            pub_log[-1]["코드3에서"]   = from_3

        df_nat_s = df_nat.loc[nat_samples].copy(); df_nat_s["표본구분"] = "전국사업체조사"
        df_pub_s = df_pub.loc[pub_samples].copy(); df_pub_s["표본구분"] = "공시대상"
        df_result = pd.concat([df_nat_s, df_pub_s], ignore_index=True)

        # ── 원본에 채택 표기 ──
        sampled_pids = set(df_result["PID"].astype(str).tolist())
        df_nat_orig  = df_nat.copy()
        df_pub_orig  = df_pub.copy()
        df_nat_orig["채택여부"] = df_nat_orig["PID"].astype(str).apply(lambda x: "채택" if x in sampled_pids else "")
        df_pub_orig["채택여부"] = df_pub_orig["PID"].astype(str).apply(lambda x: "채택" if x in sampled_pids else "")

        # ── 요약 ──
        st.markdown("### 📊 표본추출 결과")
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-blue">{len(df_result):,}</div><div class="stat-label">총 표본 수</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-green">{len(df_nat_s):,}</div><div class="stat-label">전국사업체조사</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="stat-num c-gray">{len(df_pub_s):,}</div><div class="stat-label">공시대상</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("**전국사업체조사 층별 추출 현황**")
        st.dataframe(pd.DataFrame(nat_log), use_container_width=True, hide_index=True)
        st.markdown("**공시대상 층별 추출 현황**")
        st.dataframe(pd.DataFrame(pub_log), use_container_width=True, hide_index=True)
        st.markdown("**표본 미리보기 (상위 10건)**")
        st.dataframe(df_result.head(10), use_container_width=True)

        st.markdown("### 📥 결과 다운로드")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label="✅ 표본 리스트만 다운로드",
                data=to_excel_multi({"전체표본": df_result, "전국사업체조사": df_nat_s, "공시대상": df_pub_s}),
                file_name="표본추출결과_리스트.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                label="📋 원본 명부 + 채택 표기 다운로드",
                data=to_excel_multi({"전국사업체조사_전체": df_nat_orig, "공시대상_전체": df_pub_orig}),
                file_name="표본추출결과_채택표기.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    elif not ready2:
        st.info("전국사업체조사, 공시대상, 할당표본 파일을 모두 업로드하면 추출 버튼이 활성화됩니다.")
