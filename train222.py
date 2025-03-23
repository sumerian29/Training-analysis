"""
====================================================
Employee Training & Development Analysis System
Designed by Tareq Mageed / Dhiqar Oil Co. / Ministry of Oil
====================================================
"""

import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import textwrap

# ====================================================
# Custom Styling with Background
# ====================================================
st.markdown(
    """
    <style>
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #d0f0d0 25%, #e8f8e8 50%, #f8fff8 75%);
    }
    h1, h2, h3, h4, h5 {
        color: #004d00;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ====================================================
# Company Logo
# ====================================================
st.image("sold.png", width=120)

# ====================================================
# App Title
# ====================================================
st.title("Employee Training & Development Analysis System")
st.markdown("**Analyze training data and evaluate department performance**")

# ====================================================
# Display Copyright Information in the Interface
# ====================================================
st.markdown("""
    <p style="text-align: center;">
    © Program designed and copyrighted by Tareq Mageed – Thi Qar Oil Company
    </p>
""", unsafe_allow_html=True)

# ====================================================
# Upload Excel File
# ====================================================
st.markdown("### Upload Training Data (Excel)")
uploaded_file = st.file_uploader("Select a training data Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully!")
        st.dataframe(df.head())
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()
else:
    st.warning("Please upload an Excel file to start analysis.")
    st.stop()

# ====================================================
# Performance & Efficiency Calculation
# ====================================================
required_cols = ["Training_Hours", "Exam_Score", "Attendance_Rate", "Employee_Response_Before", "Employee_Response_After"]

if all(col in df.columns for col in required_cols) and not df.empty:

    df["Performance"] = (
        (df["Exam_Score"] * 0.5) +
        (df["Attendance_Rate"] * 0.3) +
        ((df["Training_Hours"] / df["Training_Hours"].max()) * 20) +
        ((df["Employee_Response_After"] - df["Employee_Response_Before"]) * 5)
    )

    df["Employee_Efficiency"] = (
        (df["Performance"] * 0.4) +
        (df["Exam_Score"] * 0.3) +
        (df["Attendance_Rate"] * 0.3)
    )

    st.markdown("### Sample Employee Efficiency Scores")
    st.dataframe(df[["Department", "Training_Hours", "Performance", "Exam_Score", "Attendance_Rate", "Employee_Efficiency"]].head())

    # ====================================================
    # Department Ranking
    # ====================================================
    st.markdown("### Department Ranking by Efficiency")

    department_avg = df.groupby("Department")["Employee_Efficiency"].mean()
    department_ranking = department_avg.sort_values(ascending=False).reset_index()
    department_ranking.columns = ["Department", "Average Efficiency"]
    st.dataframe(department_ranking)

    best_department = department_avg.idxmax()
    worst_department = department_avg.idxmin()
    best_score = department_avg.max()
    worst_score = department_avg.min()

    if best_score != worst_score:
        st.success(f"Top Performing Department: {best_department} with average efficiency score of {best_score:.2f}")
        st.warning(f"Lowest Performing Department: {worst_department} with average efficiency score of {worst_score:.2f}")
    else:
        st.info("All departments have the same average efficiency score.")

    # ====================================================
    # Final Results Button
    # ====================================================
    if st.button("Show Final Results Only"):
        st.markdown("## Final Results")
        st.markdown("#### Department Ranking:")
        st.dataframe(department_ranking)

        if best_score != worst_score:
            st.success(f"Top Department: {best_department} - {best_score:.2f}")
            st.warning(f"Lowest Department: {worst_department} - {worst_score:.2f}")
        else:
            st.info("No performance gap detected between departments.")

    # ====================================================
    # Recommendations
    # ====================================================
    st.markdown("### Recommendations for Improvement")
    if best_score != worst_score:
        st.info(f"Consider increasing training efforts, assessment quality, or addressing specific issues in the {worst_department} department.")
    else:
        st.success("All departments are performing consistently well. No urgent action needed.")
    
    # ====================================================
    # Display Efficiency Formula
    # ====================================================
    st.markdown("### Efficiency Formula")
    st.markdown("""
    **Employee Efficiency** = (Performance * 0.4) + (Exam Score * 0.3) + (Attendance Rate * 0.3)
    
    **Performance** = Exam Score * 0.5 + Attendance Rate * 0.3 + (Training Hours / Max) * 20 + (After - Before) * 5
    """)

    # ====================================================
    # Export to PDF Button
    # ====================================================
    def create_pdf(dataframe):
        pdf_path = "department_efficiency_report.pdf"
        c = canvas.Canvas(pdf_path, pagesize=A4)
        width, height = A4

        # Margins
        margin_left = 70
        margin_top = height - 40
        line_height = 14  # height between lines

        # Header
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width / 2, margin_top, "Employee Training & Development Report")
        c.setFont("Helvetica", 11)
        c.drawCentredString(width / 2, margin_top - 20, "Thi Qar Oil Company / Ministry of Oil")

        # Department Ranking - Values Display
        c.setFont("Helvetica-Bold", 14)
        y_position = margin_top - 60
        c.drawString(margin_left, y_position, "Department Rankings by Efficiency:")

        y_position -= line_height
        for index, row in dataframe.iterrows():
            department = row['Department']
            avg_efficiency = round(row['Average Efficiency'], 2)
            c.setFont("Helvetica", 12)
            if y_position < 100:  # Check if we are near the bottom of the page
                c.showPage()  # Start a new page
                y_position = height - 40  # Reset y_position at the top of the page
            c.drawString(margin_left, y_position, f"{department}: {avg_efficiency}")
            y_position -= line_height

        # Best and Worst Performing Departments
        best_department = dataframe.loc[dataframe['Average Efficiency'].idxmax()]
        worst_department = dataframe.loc[dataframe['Average Efficiency'].idxmin()]

        c.setFont("Helvetica-Bold", 14)
        y_position -= 30  # Add some space before the best and worst departments
        c.drawString(margin_left, y_position, f"Best Performing Department: {best_department['Department']} with Efficiency: {best_department['Average Efficiency']:.2f}")
        y_position -= line_height
        c.drawString(margin_left, y_position, f"Worst Performing Department: {worst_department['Department']} with Efficiency: {worst_department['Average Efficiency']:.2f}")
        y_position -= line_height

        # Recommendations - Display as large text
        c.setFont("Helvetica-Bold", 14)
        y_position -= 20
        c.drawString(margin_left, y_position, "Recommendations for Improvement:")
        y_position -= line_height
        
        # Use textwrap to wrap long sentences
        recommendation_text = f"Consider increasing training efforts, assessment quality, or addressing specific issues in the {worst_department['Department']} department."
        wrapped_text = textwrap.wrap(recommendation_text, width=80)  # Wrap text to fit within 80 characters

        for line in wrapped_text:
            if y_position < 100:  # Check if we are near the bottom of the page
                c.showPage()  # Start a new page
                y_position = height - 40  # Reset y_position at the top of the page
            c.setFont("Helvetica", 12)
            c.drawString(margin_left, y_position, line)
            y_position -= line_height

        if best_department['Average Efficiency'] == worst_department['Average Efficiency']:
            c.setFont("Helvetica", 12)
            c.drawString(margin_left, y_position, "All departments are performing consistently well. No urgent action needed.")
            y_position -= line_height

        # Signature Area
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margin_left, 110, "Thi Qar Oil Company")
        c.drawString(margin_left, 95, "Quality Management & Institutional Development Division")
        c.line(margin_left, 70, margin_left + 150, 70)
        c.setFont("Helvetica", 9)
        c.drawString(margin_left, 55, "Signature")
        c.setFont("Helvetica-Oblique", 7)
        c.drawString(margin_left, 43, "© Program designed and copyrighted by Tareq Mageed – Thi Qar Oil Company")

        c.save()
        return pdf_path

    if st.button("Export Report to PDF"):
        path = create_pdf(department_ranking)
        with open(path, "rb") as file:
            st.download_button(label="Download PDF Report", data=file, file_name="department_efficiency_report.pdf", mime="application/pdf")
