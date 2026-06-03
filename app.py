import streamlit as st
import requests
import time
import json
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="实时礼物值计算", layout="wide")

st.title("🎁 实时礼物值计算（公网版）")

BASE_URL = "https://yapi.tuwan.com/Teacher/myGift/type/0/"

HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Referer": "https://app.tuwan.com/",
    "X-Requested-With": "XMLHttpRequest"
}

def parse_wrapped_json(text):
    text = text.strip()
    if text.startswith("(") and text.endswith(")"):
        text = text[1:-1]
    return json.loads(text)


st.markdown("### 1. 粘贴 Cookie")
cookie_text = st.text_area(
    "Cookie（浏览器复制完整 Cookie）",
    height=120,
    help="仅本次运行使用，不会保存到服务器"
)

with st.expander("📖 不会获取 Cookie？点击查看教程"):
    st.markdown("""
### 获取 Cookie
1. 登录兔玩后台（仅限网页版） https://y.tuwan.com/
2. 按 F12 打开开发者工具 ｜｜ 鼠标右键点击检查
3. 点击 Network（网络）
4. 刷新页面	
5. 点开任意请求 → Headers
6. 复制 Cookie
7. 粘贴到上方输入框
""")


st.markdown("### 2. 输入需要抓取的页数")
pages_text = st.text_input(
    "页数（例如：1,2,3,4,5）",
    value="1,2,3,4,5"
)

with st.expander("📖 不知道获取页数有什么用？点击查看教程"):
    st.markdown("""
### 获取 Cookie
1. 在礼物记录中 https://y.tuwan.com/home/gift/
2. 点击收到的礼物-> 查看赠送人
3. 赠送人送的礼物的页数即为需要实时抓取的页数
""")

check_name = st.text_input(
    "指定验算昵称（可留空）"
)

# 直接读取仓库里的文件
df_gift = pd.read_excel("gift_mapping.xlsx")

if st.button("开始计算", type="primary"):

    try:
        if not cookie_text.strip():
            st.error("请先输入 Cookie")
            st.stop()


        cookies = {}

        for item in cookie_text.split(";"):
            if "=" in item:
                k, v = item.strip().split("=", 1)
                cookies[k] = v

        pages = [
            int(x.strip())
            for x in pages_text.split(",")
            if x.strip()
        ]

        session = requests.Session()
        session.headers.update(HEADERS)
        session.cookies.update(cookies)

        all_rows = []
        progress = st.progress(0)

        for i, page in enumerate(pages):

            params = {
                "page": page,
                "step": 10,
                "_": int(time.time() * 1000)
            }

            resp = session.get(
                BASE_URL,
                params=params,
                timeout=20
            )

            data = parse_wrapped_json(resp.text)

            if data.get("error") != 0:
                st.warning(
                    f"第 {page} 页抓取失败，error={data.get('error')}"
                )
                continue

            rows = data.get("data", [])
            all_rows.extend(rows)

            progress.progress((i + 1) / len(pages))

            time.sleep(0.3)

        if not all_rows:
            st.error("没有抓取到任何数据，请检查 Cookie 是否失效")
            st.stop()

        df_log = pd.DataFrame(all_rows)[
            ["num", "create_time", "title", "name"]
        ]


        df_log["gift_name"] = (
            df_log["title"]
            .astype(str)
            .str.strip()
        )

        df_gift["title"] = (
            df_gift["title"]
            .astype(str)
            .str.strip()
        )

        df_log["create_time"] = pd.to_datetime(
            df_log["create_time"]
        )

        df_log["date"] = (
            df_log["create_time"]
            .dt.date
        )

        df_merge = df_log.merge(
            df_gift[["title", "diamond"]],
            how="left",
            left_on="gift_name",
            right_on="title"
        )

        df_merge["diamond"] = pd.to_numeric(
            df_merge["diamond"],
            errors="coerce"
        ).fillna(0)

        df_merge["num"] = pd.to_numeric(
            df_merge["num"],
            errors="coerce"
        ).fillna(0)

        df_merge["实际diamond"] = (
            df_merge["num"] *
            df_merge["diamond"]
        )

        df_sum = (
            df_merge.groupby(
                ["name", "date"],
                as_index=False
            )["实际diamond"]
            .sum()
            .rename(
                columns={
                    "实际diamond": "diamond_total"
                }
            )
            .sort_values(
                "diamond_total",
                ascending=False
            )
        )

        st.success(
            f"完成 ✔ 共 {len(df_sum)} 条汇总记录"
        )

        st.subheader("按人 + 日期汇总")

        st.dataframe(
            df_sum,
            use_container_width=True
        )

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df_sum.to_excel(
                writer,
                index=False,
                sheet_name="汇总"
            )

            df_merge.to_excel(
                writer,
                index=False,
                sheet_name="明细"
            )

        st.download_button(
            "下载统计Excel",
            data=output.getvalue(),
            file_name="礼物统计结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if check_name:

            df_check = df_merge[
                df_merge["name"] == check_name
            ].copy()

            if len(df_check):

                title_col = (
                    "title_x"
                    if "title_x" in df_check.columns
                    else "title"
                )

                df_title = (
                    df_check.groupby(
                        title_col,
                        as_index=False
                    )
                    .agg(
                        数量=("num", "sum"),
                        单价diamond=("diamond", "first"),
                        总diamond=("实际diamond", "sum")
                    )
                    .sort_values(
                        "总diamond",
                        ascending=False
                    )
                )

                st.subheader(
                    f"验算：{check_name}"
                )

                st.dataframe(
                    df_title,
                    use_container_width=True
                )

                st.metric(
                    "总diamond",
                    f"{df_title['总diamond'].sum():,.0f}"
                )

            else:
                st.info("未找到该昵称数据")

    except Exception as e:
        st.exception(e)
