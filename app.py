# Mengimpor library
import pyomo.environ as pyo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import numpy as np
import pandas as pd
import streamlit as st
import sys

# Membuat judul
st.title('OPTIMIZING ORDER SELECTION')

def preprocessing(df):
    preprocessor = ColumnTransformer([
        ('imputasi Order', SimpleImputer(strategy='constant', fill_value=0), ['Order'])],
        remainder='passthrough',
        verbose_feature_names_out=False
    )
    preprocessor.fit(df)
    df = preprocessor.transform(df)
    df = pd.DataFrame(df, columns=preprocessor.get_feature_names_out())
    return df

# Fungsi cek kolom requirement & convert data ke int
def convert_df(df):
    required_columns = ['PN', 'Order','Promise','Quality', 'Production', 'Cost', 'HPP', 'Sales']
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Missing required column: {col}")
            return

    df["Order"] = df["Order"].astype(int)
    df["Promise"] = df["Promise"].astype(int)        
    df["Quality"] = df["Quality"].astype(int)
    df["Production"] = df["Production"].astype(int)
    df["Cost"] = df["Cost"].astype(int)
    df["HPP"] = df["HPP"].astype(int)
    df["Sales"] = df["Sales"].astype(int)

# Fungsi menghitung margin
def margin(df):
    df['Margin'] = df['Sales'] - df['HPP']

# Fungsi menghitung rating
def rating(df):
    df['Rating'] = df[['Quality', 'Production', 'Cost']].dot([0.4, 0.3, 0.3])
    
# Fungsi optimasi
def solve_optimization(df,capacity):
    
    model = pyo.ConcreteModel()
    model.Pn = pyo.Var(range(len(df.PN)), bounds=(0,None))
    pn = model.Pn
    
    # Fungsi pembatas
    pn_sum = sum([pn[indeks] for indeks in range(len(df.Order))])
    model.balance = pyo.Constraint(expr = pn_sum <= capacity)

    model.limits = pyo.ConstraintList()
    for indeks in range(len(df.Order)):
        model.limits.add(expr = pn[indeks] <= df.Order[indeks])

    model.min = pyo.ConstraintList()
    for indeks in range(len(df.Order)):
        model.min.add(expr = pn[indeks] >= df.Promise[indeks])
    
    # Fungsi tujuan
    pn_sum_obj = sum([pn[indeks]*df.Rating[indeks] for indeks in range(len(df.PN))])
    model.obj = pyo.Objective(expr = pn_sum_obj, sense = maximize)

    opt = SolverFactory('glpk')
    
    results = opt.solve(model, tee=True)  # tee=True untuk menampilkan output solver di konsol
    
    # Periksa apakah solver berhasil menemukan solusi
    if results.solver.status != SolverStatus.ok or results.solver.termination_condition != TerminationCondition.optimal:
        st.error(f"Solusi tidak ditemukan! Status solver: {results.solver.status}, Termination condition: {results.solver.termination_condition}")
        return
    
    # Menambahkan garis pembatas
    st.markdown('---'*10)
    
    # Menampilkan hasil optimasi
    st.write(f'<center><b><h3>Solution: = {results.solver.termination_condition} </b></h3>', unsafe_allow_html=True)
    margin = []
    for i in range(len(pn)):
        part_value = pyo.value(pn[i])
        value =  part_value * df.Margin[i]
        margin.append(value)
        if pyo.value(pn[i]) > 0:
            st.write(f'<center><b><h3>Part Number: {df.PN[i]} = {pyo.value(pn[i]):,.0f} pcs</b></h3>', unsafe_allow_html=True)
            
    total_margin = sum(margin)
    st.write(f'<center><b><h3>Total Margin: {total_margin:,.0f} </b></h3>', unsafe_allow_html=True)

    # Membuat DataFrame untuk hasil optimasi yang ingin diunduh
    result_df = pd.DataFrame({
        'Part Number': df.PN,
        'Quantity (pcs)': [pyo.value(pn[i]) for i in range(len(pn))],
        'Margin Value': margin
    })

    # Menghapus baris dengan Quantity (pcs) = 0
    result_df = result_df[result_df['Quantity (pcs)'] > 0]

    # Menambahkan total margin ke dalam DataFrame
    result_df.loc[len(result_df)] = ['Total Margin', None, total_margin]

    # Menyimpan file Excel di disk sementara
    file_name = "optimized_results.xlsx"
    result_df.to_excel(file_name, index=False, sheet_name='Optimization Results')
    
    # Membuat tombol download untuk file Excel
    st.download_button(
        label="Download Optimized Results (Excel)",
        data=open(file_name, 'rb').read(),  # Membaca file dan mengirimkan sebagai binary data
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel Master Data", type=["xlsx"])

# Upload excel
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        preprocessing(df)
        convert_df(df)
        margin(df)
        rating(df)
        st.write(f'Bobot -> Quality: 40%, Production: 30%, Cost: 30%')
        st.write(df)  
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        sys.exit()
        
    # Input capacity
    capacity = st.number_input("Enter Capacity:", min_value=0)

    # Tombol ekskusi optimasi
    if st.button("Calculate"):
        #try:
        solve_optimization(df,capacity)
        #except Exception as e:
        #st.error(f"Error : {e}")

