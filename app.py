import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excelプレビュー (Streamlit)", page_icon="📄", layout="wide")

st.title("📄 Excelプレビュー")
st.caption("ファイルをアップロードして、シートや範囲を選ぶだけでプレビューできます。CSVもOK。")

with st.sidebar:
    st.header("アップロード")
    file = st.file_uploader("Excel/CSVファイルを選択", type=["xlsx","xls","xlsb","csv"])
    st.markdown("---")
    st.subheader("読み取りオプション")
    header_row = st.number_input("ヘッダー行（0=ヘッダーなし）", min_value=0, value=1, step=1,
                                 help="1以上の場合、その行を列名として扱います。0の場合は自動で列名を割り当てます。")
    usecols = st.text_input("列範囲（例: A:D または A,C,E）", value="", help="空欄で全列。Excel形式の指定。CSVでは無視されます。")
    nrows = st.number_input("最大表示行数", min_value=1, value=500, step=50)
    skiprows = st.text_input("先頭スキップ行数/リスト", value="0", help="例: 1 または 1,2,5")
    parse_dates = st.checkbox("日付自動解析（CSV）", value=True)
    st.markdown("---")
    st.caption("ヒント: 大きなファイルは最大行数を減らすと軽くなります。")

def _parse_skiprows(s):
    s = (s or "").strip()
    if not s:
        return None
    try:
        if "," in s:
            return [int(x.strip()) for x in s.split(",") if x.strip()]
        return int(s)
    except Exception:
        return None

@st.cache_data(show_spinner=False)
def load_csv(bytes_data, header_row, nrows, skiprows, parse_dates):
    header = None if header_row == 0 else header_row - 1
    return pd.read_csv(io.BytesIO(bytes_data),
                       header=header,
                       nrows=nrows if nrows else None,
                       skiprows=_parse_skiprows(skiprows),
                       parse_dates=parse_dates)

@st.cache_data(show_spinner=False)
def load_excel(bytes_data, sheet_name, header_row, usecols, nrows, skiprows, engine):
    header = None if header_row == 0 else header_row - 1
    return pd.read_excel(io.BytesIO(bytes_data),
                         sheet_name=sheet_name,
                         engine=engine,
                         header=header,
                         usecols=(usecols or None),
                         nrows=nrows if nrows else None,
                         skiprows=_parse_skiprows(skiprows))

def detect_engine(xls):
    # Decide engine by extension
    name = (xls.name or "").lower()
    if name.endswith(".xlsx"):
        return "openpyxl"
    if name.endswith(".xls"):
        return "xlrd"
    if name.endswith(".xlsb"):
        return "pyxlsb"
    return None

if not file:
    st.info("左のサイドバーから Excel/CSV ファイルをアップロードしてください。")
    st.stop()

bytes_data = file.getvalue()
name = file.name.lower()

if name.endswith(".csv"):
    df = load_csv(bytes_data, header_row, nrows, skiprows, parse_dates)
    st.success(f"CSV を読み込みました：**{file.name}**  /  形状: {df.shape[0]}行 × {df.shape[1]}列")
    st.dataframe(df, use_container_width=True, height=520)
else:
    engine = detect_engine(file)
    try:
        # List sheets
        xls = pd.ExcelFile(io.BytesIO(bytes_data), engine=engine)
    except Exception as e:
        st.error(f"Excelファイルの解析に失敗しました: {e}")
        st.stop()

    with st.sidebar:
        sheet = st.selectbox("シートを選択", xls.sheet_names, index=0)

    try:
        df = load_excel(bytes_data, sheet, header_row, usecols, nrows, skiprows, engine)
    except ValueError as ve:
        st.error(f"列範囲(usecols)が不正かもしれません: {ve}")
        st.stop()
    except Exception as e:
        st.error(f"読み込みに失敗しました: {e}")
        st.stop()

    st.success(f"Excel を読み込みました：**{file.name}** / シート: **{sheet}** / 形状: {df.shape[0]}行 × {df.shape[1]}列")
    st.dataframe(df, use_container_width=True, height=520)

with st.expander("メタ情報 / 列型を確認"):
    st.write("列型:")
    st.write(df.dtypes.astype(str))
    buf = io.StringIO()
    df.info(buf=buf)
    st.text(buf.getvalue())

# Download preview
csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "このプレビューをCSVでダウンロード",
    data=csv_bytes,
    file_name="preview.csv",
    mime="text/csv",
    use_container_width=True
)

st.caption("Powered by Streamlit + pandas | 使い方: 左のサイドバーでオプションを調整できます。")
