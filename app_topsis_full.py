
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="TOPSIS - MCDM App", layout="wide")
st.title("TOPSIS (Technique for Order Preference by Similarity to Ideal Solution)")

uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])

def normalize(matrix):
    norm = np.sqrt((matrix**2).sum(axis=0))
    return matrix / norm

def calculate_topsis(decision_matrix, weights, types, criteria, alternatives):
    norm_matrix = normalize(decision_matrix)
    weighted_matrix = norm_matrix * weights

    ideal = []
    anti_ideal = []
    for i, t in enumerate(types):
        if t.lower() == "benefit":
            ideal.append(np.max(weighted_matrix[:, i]))
            anti_ideal.append(np.min(weighted_matrix[:, i]))
        else:
            ideal.append(np.min(weighted_matrix[:, i]))
            anti_ideal.append(np.max(weighted_matrix[:, i]))

    d_pos = np.sqrt(((weighted_matrix - ideal)**2).sum(axis=1))
    d_neg = np.sqrt(((weighted_matrix - anti_ideal)**2).sum(axis=1))
    scores = d_neg / (d_pos + d_neg)

    result_df = pd.DataFrame({
        "TOPSIS Score": scores,
        "Rank": scores.argsort()[::-1].argsort() + 1
    }, index=alternatives)

    ideal_df = pd.DataFrame({
        "Criterion": criteria,
        "Ideal": ideal,
        "Anti-Ideal": anti_ideal
    })

    return norm_matrix, weighted_matrix, ideal_df, result_df

def generate_figure(decision_matrix):
    fig, ax = plt.subplots(figsize=(6, 2))
    ax.axis("off")
    table_data = [["Alternative"] + list(decision_matrix.columns)]
    for alt, row in decision_matrix.iterrows():
        table_data.append([alt] + list(row))
    table = ax.table(cellText=table_data, loc="center", cellLoc="center")
    buf = BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    return buf

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    decision_df = pd.read_excel(xls, sheet_name="DecisionMatrix", index_col=0)
    weights_df = pd.read_excel(xls, sheet_name="Weights")
    types_df = pd.read_excel(xls, sheet_name="Types")

    criteria = decision_df.columns.tolist()
    alternatives = decision_df.index.tolist()
    weights = weights_df.iloc[0].values
    types = types_df.iloc[0].values
    decision_matrix = decision_df.values

    norm_matrix, weighted_matrix, ideal_df, result_df = calculate_topsis(
        decision_matrix, weights, types, criteria, alternatives
    )

    st.subheader("📌 Step 1: Raw Decision Matrix")
    st.dataframe(decision_df)

    st.subheader("📌 Step 2: Normalized Matrix")
    st.dataframe(pd.DataFrame(norm_matrix, index=alternatives, columns=criteria))

    st.subheader("📌 Step 3: Weighted Normalized Matrix")
    st.dataframe(pd.DataFrame(weighted_matrix, index=alternatives, columns=criteria))

    st.subheader("📌 Step 4: Ideal & Anti-Ideal Solutions")
    st.dataframe(ideal_df)

    st.subheader("📌 Step 5: Final TOPSIS Ranking")
    st.dataframe(result_df.style.highlight_max(axis=0, color="lightgreen"))

    st.subheader("🖼️ Raw Matrix Table Figure")
    fig_buf = generate_figure(decision_df)
    st.image(fig_buf, caption="Raw Decision Matrix", use_column_width=True)
    st.download_button("📥 Download Table Figure (PNG)", data=fig_buf, file_name="topsis_raw_matrix.png", mime="image/png")

    def to_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            decision_df.to_excel(writer, sheet_name="1_RawData")
            pd.DataFrame(norm_matrix, index=alternatives, columns=criteria).to_excel(writer, sheet_name="2_Normalized")
            pd.DataFrame(weighted_matrix, index=alternatives, columns=criteria).to_excel(writer, sheet_name="3_Weighted")
            ideal_df.to_excel(writer, sheet_name="4_Ideal_AntiIdeal", index=False)
            result_df.to_excel(writer, sheet_name="5_Results")
        output.seek(0)
        return output

    st.download_button("📥 Download Full Report (Excel)", data=to_excel(), file_name="TOPSIS_Report.xlsx")
