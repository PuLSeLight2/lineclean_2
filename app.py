import streamlit as st

st.set_page_config(
    page_title="Line Clean Wastage",
    page_icon="🍺",
    layout="centered",
    initial_sidebar_state="collapsed",
)

PRODUCTS = [
    "Coors",
    "Madri",
    "Stowford",
    "Carling",
    "Carling Darkfruits",
    "Guinness",
    "Alpaca",
]

PINTS_PER_TAP = 5

st.markdown("""
<style>
    .block-container { max-width: 760px; padding-top: 1.4rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] {
        background: rgba(128,128,128,0.09);
        border: 1px solid rgba(128,128,128,0.18);
        padding: 14px 16px;
        border-radius: 14px;
    }
    .small-note { opacity: .75; font-size: .9rem; }
    div.stButton > button { width: 100%; border-radius: 12px; min-height: 2.8rem; }
</style>
""", unsafe_allow_html=True)

st.title("🍺 Line Clean Wastage")
st.caption("Record the taps used in each pull-through. Each tap = 5 pints of wastage.")

if "first" not in st.session_state:
    st.session_state.first = {p: 0 for p in PRODUCTS}
if "second" not in st.session_state:
    st.session_state.second = {p: 0 for p in PRODUCTS}
if "quick_product" not in st.session_state:
    st.session_state.quick_product = "Alpaca"

with st.expander("Quick entry — add taps to a product", expanded=True):
    product = st.selectbox(
        "Product",
        PRODUCTS,
        index=PRODUCTS.index(st.session_state.quick_product),
        key="quick_product",
    )
    pull = st.radio(
        "Which pull-through?",
        ["First pull-through", "Second pull-through"],
        horizontal=True,
        key="quick_pull",
    )
    taps = st.number_input(
        "Number of taps",
        min_value=0,
        max_value=100,
        value=0,
        step=1,
        key="quick_taps",
    )
    if st.button("Add taps", type="primary"):
        target = st.session_state.first if pull == "First pull-through" else st.session_state.second
        target[product] += int(taps)
        st.success(f"Added {int(taps)} tap(s) to {product} — {pull.lower()}.")

st.divider()
st.subheader("1. First pull-through")
st.caption("Enter how many taps you used for each product.")
first_cols = st.columns(2)
for i, p in enumerate(PRODUCTS):
    with first_cols[i % 2]:
        st.number_input(
            f"{p} taps",
            min_value=0,
            max_value=100,
            step=1,
            value=int(st.session_state.first[p]),
            key=f"first_{p}",
            on_change=lambda product=p: st.session_state.first.update(
                {product: int(st.session_state[f"first_{product}"])}
            ),
        )

st.divider()
st.subheader("2. Second pull-through")
st.caption("Enter the tap counts again for the second pull-through.")
second_cols = st.columns(2)
for i, p in enumerate(PRODUCTS):
    with second_cols[i % 2]:
        st.number_input(
            f"{p} taps",
            min_value=0,
            max_value=100,
            step=1,
            value=int(st.session_state.second[p]),
            key=f"second_{p}",
            on_change=lambda product=p: st.session_state.second.update(
                {product: int(st.session_state[f"second_{product}"])}
            ),
        )

# Read current widget values to keep totals live and accurate.
first_counts = {p: int(st.session_state.get(f"first_{p}", st.session_state.first[p])) for p in PRODUCTS}
second_counts = {p: int(st.session_state.get(f"second_{p}", st.session_state.second[p])) for p in PRODUCTS}
for p in PRODUCTS:
    st.session_state.first[p] = first_counts[p]
    st.session_state.second[p] = second_counts[p]

first_taps = sum(first_counts.values())
second_taps = sum(second_counts.values())
first_waste = first_taps * PINTS_PER_TAP
second_waste = second_taps * PINTS_PER_TAP
total_taps = first_taps + second_taps
total_waste = first_waste + second_waste

st.divider()
st.subheader("Wastage summary")
m1, m2, m3 = st.columns(3)
m1.metric("First pull-through", f"{first_waste} pints", f"{first_taps} taps")
m2.metric("Second pull-through", f"{second_waste} pints", f"{second_taps} taps")
m3.metric("Total wastage", f"{total_waste} pints", f"{total_taps} taps")

st.markdown("#### Breakdown by product")
rows = []
for p in PRODUCTS:
    a, b = first_counts[p], second_counts[p]
    if a or b:
        rows.append({
            "Product": p,
            "First pull-through": f"{a} taps · {a * PINTS_PER_TAP} pints",
            "Second pull-through": f"{b} taps · {b * PINTS_PER_TAP} pints",
            "Total wastage": f"{(a + b) * PINTS_PER_TAP} pints",
        })

if rows:
    st.table(rows)
else:
    st.info("Enter tap counts above to see your wastage breakdown.")

st.caption("Calculation: taps used × 5 pints. Counts are kept in this app session; reset the page to start a fresh line clean.")

if st.button("Reset all tap counts"):
    for p in PRODUCTS:
        st.session_state.first[p] = 0
        st.session_state.second[p] = 0
        st.session_state[f"first_{p}"] = 0
        st.session_state[f"second_{p}"] = 0
    st.rerun()
