
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="TOPSIS - MCDM App", layout="wide")
st.title("TOPSIS (Technique for Order Preference by Similarity to Ideal Solution)")

uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])

def normalize(matrix):
    norm = np.sqrt((matrix**2).sum(axis=0))
    return matrix / norm

def calculate_topsis(decision_matrix, weights, types):
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
    return scores, weighted_matrix, ideal, anti_ideal, norm_matrix

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    decision_matrix = pd.read_excel(xls, sheet_name="DecisionMatrix", index_col=0)
    weights_df = pd.read_excel(xls, sheet_name="Weights")
    types_df = pd.read_excel(xls, sheet_name="Types")

    weights = weights_df.iloc[0].values  # horizontal row of weights
    types = types_df.iloc[0].values      # horizontal row of types

    scores, weighted_matrix, ideal, anti_ideal, norm_matrix = calculate_topsis(
        decision_matrix.values, weights, types
    )

    result_df = pd.DataFrame({
        "Alternative": decision_matrix.index,
        "TOPSIS Score": scores,
        "Rank": scores.argsort()[::-1].argsort() + 1
    }).set_index("Alternative")

    st.subheader("📈 Ranking Results")
    st.dataframe(result_df.style.highlight_max(axis=0, color="lightgreen"))

    st.subheader("📊 Weighted Normalized Matrix")
    st.dataframe(pd.DataFrame(weighted_matrix, index=decision_matrix.index, columns=decision_matrix.columns))

    st.subheader("🏆 Ideal & Anti-Ideal Solutions")
    st.write("Ideal Solution:", ideal)
    st.write("Anti-Ideal Solution:", anti_ideal)

    # Download as Excel
    def to_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            decision_matrix.to_excel(writer, sheet_name="DecisionMatrix")
            weights_df.to_excel(writer, sheet_name="Weights", index=False)
            types_df.to_excel(writer, sheet_name="Types", index=False)
            pd.DataFrame(norm_matrix, index=decision_matrix.index, columns=decision_matrix.columns).to_excel(writer, sheet_name="Normalized")
            pd.DataFrame(weighted_matrix, index=decision_matrix.index, columns=decision_matrix.columns).to_excel(writer, sheet_name="Weighted")
            result_df.to_excel(writer, sheet_name="Results")
        output.seek(0)
        return output

    st.download_button("📥 Download Full Results (Excel)", data=to_excel(), file_name="TOPSIS_Results.xlsx")
