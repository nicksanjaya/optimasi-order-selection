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

# Fungsi cek kolom requirement & convert data ke int
def convert_df(df):
    required_columns = ['PN', 'Quality', 'Production', 'Cost', 'HPP', 'Sales']
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Missing required column: {col}")
            return
            
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

#fungsi input order
def input_order(df):
    pn_values = {}
    for part in df.PN:
        pn_values[part] = st.number_input(f"Enter Quantity Part Number {part}:", min_value=0)
    data = {
    'PN': df.PN,
    'Qty': [pn_values[part] for part in df.PN]
    }
    global order
    order  = pd.DataFrame(data)

# Fungsi optimasi
def solve_optimization(df,order,capacity):
    
    model = pyo.ConcreteModel()
    model.Pn = pyo.Var(range(len(df.PN)), bounds=(0,None))
    pn = model.Pn
    
    # Fungsi pembatas
    pn_sum = sum([pn[indeks] for indeks in range(len(order.PN))])
    model.balance = pyo.Constraint(expr = pn_sum <= capacity)

    model.limits = pyo.ConstraintList()
    for indeks in range(len(order.PN)):
        model.limits.add(expr = pn[indeks] <= order.Qty[indeks])

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
    margin = []
    for i in range(len(pn)):
        part_value = pyo.value(pn[i])
        value =  part_value * df.Margin[i]
        margin.append(value)
        if pyo.value(pn[i]) > 0:
            st.write(f'<center><b><h3>Part Number: {df.PN[i]} = {pyo.value(pn[i]):,.0f}</b></h3>', unsafe_allow_html=True)
            
    total_margin = sum(margin)
    st.write(f'<center><b><h3>Total Margin: = {total_margin:,.0f} </b></h3>', unsafe_allow_html=True)

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel Master Data", type=["xlsx"])

# Upload excel
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
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
    
    input_order(df)

    # Tombol ekskusi optimasi
    if st.button("Calculate"):
        try:
            solve_optimization(df,order,capacity)
        except Exception as e:
            st.error(f"Error : {e}")
