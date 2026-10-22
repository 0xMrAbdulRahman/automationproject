import streamlit as st

import part1
import part2

st.set_page_config(page_title="التقرير الخارجي", layout="wide")

if "page" not in st.session_state:
    st.session_state["page"] = "home"


def go(page: str):
    st.session_state["page"] = page


page = st.session_state["page"]

if page == "home":
    st.title("التقرير الخارجي")
    st.write("اختار الجزء اللي عايز تشتغل عليه:")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("الجزء الأول")
        st.caption("تقرير المحفظة: التليفونات، الهويات المشتركة، الحالة، المحصل، المديونية، وشيتين آخر يومين.")
        st.button("افتح الجزء الأول", on_click=go, args=("part1",),
                  type="primary", use_container_width=True)
    with c2:
        st.subheader("الجزء الثاني")
        st.caption("توزيع أعداد العملاء على المحصلين واختيار العملاء وتجهيز التقرير النهائي.")
        st.button("افتح الجزء الثاني", on_click=go, args=("part2",),
                  type="primary", use_container_width=True)
else:
    st.button("← الرجوع للـ Home", on_click=go, args=("home",))
    if page == "part1":
        part1.render()
    else:
        part2.render()