# Mengimpor library
import pyomo.environ as pyo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import numpy as np
import pandas as pd
import streamlit as st


# Membuat judul
st.title('OPTIMIZATION ORDER')

# Menambah subheader
st.subheader('Selamat datang di aplikasi optimasi kuota order ke vendor subcon')

def upload_data():
    df = pd.read_excel('Part Number Order.xlsx')
    return df

def convert_df(df):
    # Check if necessary columns exist
    required_columns = ['PN', 'Quality', 'Pas', 'Cost']
    for col in required_columns:
        if col not in df.columns:
            st.error(f"Missing required column: {col}")
            return
            
    df["Quality"] = df["Quality"].astype(int)
    df["Pas"] = df["Pas"].astype(int)
    df["Cost"] = df["Cost"].astype(int)

def grading(df):
    df['Grade'] = df[['Quality', 'Pas', 'Cost']].dot([0.4, 0.3, 0.3])

def solve_optimization(df,order,capacity):
    
    model = pyo.ConcreteModel()
    model.Pn = pyo.Var(range(len(df.PN)), bounds=(0,None))
    pn = model.Pn
    
    # Fungsi pembatas
    pn_sum = sum([pn[indeks] for indeks in range(len(order.PN))])
    model.balance = pyo.Constraint(expr = pn_sum == capacity)

    model.limits = pyo.ConstraintList()
    for indeks in range(len(order.PN)):
        model.limits.add(expr = pn[indeks] <= order.Qty[indeks])

    # Fungsi tujuan
    pn_sum_obj = sum([pn[indeks]*df.Grade[indeks] for indeks in range(len(df.PN))])
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
    for i in range(len(pn)):
        st.write('<center><b><h3>Part Number: ', df.PN[i], '=', pyo.value(pn[i]), '</b></h3>', unsafe_allow_html=True)


# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel Master Data", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        convert_df(df)
        grading(df)
        st.write(df)  
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        
    # Input box for capacity
    capacity = st.number_input("Enter Capacity:", min_value=0)
    pn_abc = st.number_input("Enter Part Number ABC:", min_value=0)
    pn_def = st.number_input("Enter Part Number DEF:", min_value=0)
    pn_ghi = st.number_input("Enter Part Number GHI:", min_value=0)
    pn_jkl = st.number_input("Enter Part Number JKL:", min_value=0)
    pn_mno = st.number_input("Enter Part Number MNO:", min_value=0)
    data = {
    'PN' : ['ABC', 'DEF', 'GHI', 'JKL', 'MNO'],
    'Qty' : [pn_abc, pn_def, pn_ghi, pn_jkl, pn_mno]
    }
    
    order  = pd.DataFrame(data)


    # Button to create schedule
    if st.button("Calculate"):
        try:
            solve_optimization(df,order,capacity)
        except Exception as e:
            st.error(f"Error : {e}")