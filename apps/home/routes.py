# -*- encoding: utf-8 -*-
"""
Copyright (c) 2019 - present AppSeed.us
"""

import wtforms
from apps.home import blueprint
from flask import render_template, request, redirect, url_for
from flask_login import login_required
from jinja2 import TemplateNotFound
from flask_login import login_required, current_user
from apps import db, config
from apps.models import *
from apps.tasks import *
from apps.authentication.models import Users
from flask_wtf import FlaskForm
import pandas as pd
import plotly
import plotly.express as px
import plotly.graph_objects as go
import json
import os

@blueprint.route('/')
@blueprint.route('/index')
def index():
    """
    Fungsi utama untuk merender halaman dashboard.
    Membaca semua sheet relevan dari file Excel, mengubahnya menjadi tabel HTML,
    dan membuat visualisasi data dengan Plotly.
    """
    try:
        # --- Lokasi File Excel ---
        file_path = 'Analisis Dampak Ekspor-Impor Pendekatan Tahun 2020-2025 (Covered).xlsx'

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File Excel tidak ditemukan di lokasi: {file_path}")

        xl = pd.ExcelFile(file_path)

        # --- Daftar Sheet yang akan dimuat ---
        sheets_to_load = {
            'Periode & Neraca Perdagangan': 0,
            'Hasil Analisis 2020-2025': 0,
            '2020': 17,
            '2021': 17,
            '2022': 17,
            '2023': 17,
            '2024': 17,
            '2025': 42,
            'Komoditas dan Agregasi': 2,
            'FOB Negara (Ekspor-Impor)': 2,
        }

        # --- Memuat DataFrames dan Mengonversinya ke Tabel HTML ---
        tables = {}
        dataframes = {}
        for sheet_name, skip_rows in sheets_to_load.items():
            try:
                if sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name, skiprows=skip_rows)
                    df.dropna(axis=1, how='all', inplace=True)
                    df.dropna(axis=0, how='all', inplace=True)
                    df.reset_index(drop=True, inplace=True)
                    dataframes[sheet_name] = df
                    tables[sheet_name] = df.to_html(
                        classes="table table-striped table-bordered datatable-class", 
                        index=False, 
                        table_id=f"table-{sheet_name.replace(' ', '-').replace('&', '')}"
                    )
                    print(f"Sheet '{sheet_name}' berhasil dimuat dengan {len(df)} baris")
                    print(f"Kolom: {list(df.columns)}")
                else:
                    tables[sheet_name] = f"<p class='text-danger'>Sheet '{sheet_name}' tidak ditemukan.</p>"
            except Exception as e:
                tables[sheet_name] = f"<p class='text-danger'>Gagal memuat sheet '{sheet_name}': {e}</p>"
                dataframes[sheet_name] = pd.DataFrame()

        # --- Membuat Grafik Interaktif dengan Plotly ---
        graphs = {}
        
        # Debugging: Print available dataframes
        print("Available dataframes:", list(dataframes.keys()))
        
        df_periode = dataframes.get('Periode & Neraca Perdagangan')
        if df_periode is not None and not df_periode.empty:
            print(f"DataFrame Periode shape: {df_periode.shape}")
            print(f"Kolom DataFrame Periode: {list(df_periode.columns)}")
            print(f"Sample data:\n{df_periode.head()}")
            
            # Coba berbagai kemungkinan nama kolom
            possible_period_cols = ['Period', 'Periode', 'Tahun', 'Year', 'Date', 'Bulan']
            possible_export_cols = ['Total_Export', 'Ekspor', 'Export', 'Total Ekspor', 'Nilai Ekspor']
            possible_import_cols = ['Total_Import', 'Impor', 'Import', 'Total Impor', 'Nilai Impor']
            possible_balance_cols = ['Trade_Balance', 'Neraca', 'Balance', 'Neraca Perdagangan', 'Saldo']
            
            period_col = None
            export_col = None
            import_col = None
            balance_col = None
            
            # Cari kolom yang sesuai
            for col in df_periode.columns:
                col_lower = str(col).lower()
                if any(pc.lower() in col_lower for pc in possible_period_cols):
                    period_col = col
                elif any(ec.lower() in col_lower for ec in possible_export_cols):
                    export_col = col
                elif any(ic.lower() in col_lower for ic in possible_import_cols):
                    import_col = col
                elif any(bc.lower() in col_lower for bc in possible_balance_cols):
                    balance_col = col
            
            print(f"Detected columns - Period: {period_col}, Export: {export_col}, Import: {import_col}, Balance: {balance_col}")
            
            if period_col and (export_col or import_col):
                try:
                    # Buat copy dataframe untuk manipulasi
                    df_work = df_periode.copy()
                    
                    # Konversi kolom numerik
                    numeric_cols = []
                    if export_col:
                        df_work[export_col] = pd.to_numeric(df_work[export_col], errors='coerce')
                        numeric_cols.append(export_col)
                    if import_col:
                        df_work[import_col] = pd.to_numeric(df_work[import_col], errors='coerce')
                        numeric_cols.append(import_col)
                    if balance_col:
                        df_work[balance_col] = pd.to_numeric(df_work[balance_col], errors='coerce')
                        numeric_cols.append(balance_col)
                    
                    # Hapus baris dengan nilai NaN
                    df_work.dropna(subset=numeric_cols, inplace=True)
                    
                    # Coba konversi periode ke datetime atau ekstrak tahun
                    try:
                        # Coba format datetime
                        df_work[period_col] = pd.to_datetime(df_work[period_col], errors='coerce')
                        df_work['Year'] = df_work[period_col].dt.year
                    except:
                        try:
                            # Coba ekstrak tahun langsung
                            df_work['Year'] = pd.to_numeric(df_work[period_col], errors='coerce')
                        except:
                            # Fallback: gunakan index sebagai tahun
                            df_work['Year'] = range(2020, 2020 + len(df_work))
                    
                    # Agregasi per tahun jika ada data bulanan
                    if len(df_work) > 6:  # Jika lebih dari 6 data point, kemungkinan data bulanan
                        df_yearly = df_work.groupby('Year')[numeric_cols].sum().reset_index()
                    else:
                        df_yearly = df_work.copy()
                        if 'Year' not in df_yearly.columns:
                            df_yearly['Year'] = df_yearly[period_col]
                    
                    print(f"Data untuk grafik:\n{df_yearly}")
                    
                    # Grafik 1: Tren Tahunan Ekspor dan Impor
                    if export_col and import_col:
                        fig1 = go.Figure()
                        fig1.add_trace(go.Scatter(
                            x=df_yearly['Year'], 
                            y=df_yearly[export_col],
                            mode='lines+markers',
                            name='Ekspor',
                            line=dict(color='#2E8B57', width=3),
                            marker=dict(size=8)
                        ))
                        fig1.add_trace(go.Scatter(
                            x=df_yearly['Year'], 
                            y=df_yearly[import_col],
                            mode='lines+markers',
                            name='Impor',
                            line=dict(color='#DC143C', width=3),
                            marker=dict(size=8)
                        ))
                        fig1.update_layout(
                            title='Tren Tahunan Ekspor dan Impor (2020-2025)',
                            xaxis_title='Tahun',
                            yaxis_title='Total Nilai (Juta US$)',
                            template='plotly_white',
                            height=400
                        )
                        graphs['tren_tahunan'] = json.dumps(fig1, cls=plotly.utils.PlotlyJSONEncoder)
                    
                    # Grafik 2: Neraca Perdagangan
                    if balance_col:
                        colors = ['#28a745' if val >= 0 else '#dc3545' for val in df_yearly[balance_col]]
                        fig2 = go.Figure(data=[
                            go.Bar(
                                x=df_yearly['Year'], 
                                y=df_yearly[balance_col],
                                marker_color=colors
                            )
                        ])
                        fig2.update_layout(
                            title='Neraca Perdagangan Tahunan (2020-2025)',
                            xaxis_title='Tahun',
                            yaxis_title='Total Nilai (Juta US$)',
                            template='plotly_white',
                            height=400
                        )
                        graphs['neraca_tahunan'] = json.dumps(fig2, cls=plotly.utils.PlotlyJSONEncoder)
                    elif export_col and import_col:
                        # Hitung neraca jika tidak ada kolom balance
                        df_yearly['calculated_balance'] = df_yearly[export_col] - df_yearly[import_col]
                        colors = ['#28a745' if val >= 0 else '#dc3545' for val in df_yearly['calculated_balance']]
                        fig2 = go.Figure(data=[
                            go.Bar(
                                x=df_yearly['Year'], 
                                y=df_yearly['calculated_balance'],
                                marker_color=colors
                            )
                        ])
                        fig2.update_layout(
                            title='Neraca Perdagangan Tahunan (2020-2025)',
                            xaxis_title='Tahun',
                            yaxis_title='Total Nilai (Juta US$)',
                            template='plotly_white',
                            height=400
                        )
                        graphs['neraca_tahunan'] = json.dumps(fig2, cls=plotly.utils.PlotlyJSONEncoder)
                    
                except Exception as e:
                    print(f"Error creating period graphs: {e}")
            else:
                print("Kolom yang diperlukan tidak ditemukan untuk grafik periode")

        # Grafik 3: Dari sheet 'FOB Negara (Ekspor-Impor)'
        df_negara = dataframes.get('FOB Negara (Ekspor-Impor)')
        if df_negara is not None and not df_negara.empty:
            print(f"DataFrame Negara shape: {df_negara.shape}")
            print(f"Kolom DataFrame Negara: {list(df_negara.columns)}")
            
            # Cari kolom negara dan tahun 2024
            negara_col = None
            value_col = None
            
            for col in df_negara.columns:
                col_str = str(col).lower()
                if 'negara' in col_str or 'country' in col_str or 'tujuan' in col_str:
                    negara_col = col
                elif '2024' in str(col):
                    value_col = col
            
            print(f"Detected columns - Negara: {negara_col}, Value: {value_col}")
            
            if negara_col and value_col:
                try:
                    df_negara_work = df_negara[[negara_col, value_col]].copy()
                    df_negara_work[value_col] = pd.to_numeric(df_negara_work[value_col], errors='coerce')
                    df_negara_work.dropna(inplace=True)
                    df_negara_sorted = df_negara_work.sort_values(by=value_col, ascending=False).head(15)
                    
                    fig3 = go.Figure(data=[
                        go.Bar(
                            x=df_negara_sorted[negara_col], 
                            y=df_negara_sorted[value_col],
                            marker_color='#1f77b4'
                        )
                    ])
                    fig3.update_layout(
                        title='Top 15 Negara Tujuan Ekspor (FOB) Tahun 2024',
                        xaxis_title='Negara',
                        yaxis_title='Nilai FOB (Juta US$)',
                        template='plotly_white',
                        height=500,
                        xaxis={'tickangle': 45}
                    )
                    graphs['ekspor_impor_negara'] = json.dumps(fig3, cls=plotly.utils.PlotlyJSONEncoder)
                except Exception as e:
                    print(f"Error creating country graph: {e}")

        print(f"Generated graphs: {list(graphs.keys())}")

        return render_template(
            'pages/ekspor_impor.html',
            tables=tables,
            graphs=graphs,
            sheet_names=list(sheets_to_load.keys())
        )

    except FileNotFoundError as e:
        return render_template('pages/error.html', error_message=str(e))
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return render_template('pages/error.html', error_message=f"Terjadi kesalahan yang tidak terduga: {e}")


# ... (sisa kode tidak berubah) ...

@blueprint.route('/tables')
def tables():
    context = {
        'segment': 'tables'
    }
    return render_template('pages/tables.html', **context)

@blueprint.route('/billing')
def billing():
    context = {
        'segment': 'billing'
    }
    return render_template('pages/billing.html', **context)

@blueprint.route('/virtual-reality')
def virtual_reality():
    context = {
        'segment': 'virtual_reality'
    }
    return render_template('pages/virtual-reality.html', **context)

@blueprint.route('/rtl')
def rtl():
    context = {
        'segment': 'rtl'
    }
    return render_template('pages/rtl.html', **context)

@blueprint.route('/notifications')
def notifications():
    context = {
        'segment': 'notifications'
    }
    return render_template('pages/notifications.html', **context)

@blueprint.route('/icons')
def icons():
    context = {
        'segment': 'icons'
    }
    return render_template('pages/icons.html', **context)

@blueprint.route('/map')
def map():
    context = {
        'segment': 'map'
    }
    return render_template('pages/map.html', **context)

@blueprint.route('/typography')
def typography():
    context = {
        'segment': 'typography'
    }
    return render_template('pages/typography.html', **context)

@blueprint.route('/template')
def template():
    context = {
        'segment': 'template'
    }
    return render_template('pages/template.html', **context)

@blueprint.route('/landing')
def landing():
    context = {
        'segment': 'landing'
    }
    return render_template('pages/landing.html', **context)

def getField(column): 
    if isinstance(column.type, db.Text):
        return wtforms.TextAreaField(column.name.title())
    if isinstance(column.type, db.String):
        return wtforms.StringField(column.name.title())
    if isinstance(column.type, db.Boolean):
        return wtforms.BooleanField(column.name.title())
    if isinstance(column.type, db.Integer):
        return wtforms.IntegerField(column.name.title())
    if isinstance(column.type, db.Float):
        return wtforms.DecimalField(column.name.title())
    if isinstance(column.type, db.LargeBinary):
        return wtforms.HiddenField(column.name.title())
    return wtforms.StringField(column.name.title()) 

@blueprint.route('/profile', methods=['GET', 'POST'])
@login_required
def profile():

    class ProfileForm(FlaskForm):
        pass

    readonly_fields = Users.readonly_fields
    full_width_fields = {"bio"}

    for column in Users.__table__.columns:
        if column.name == "id":
            continue

        field_name = column.name
        if field_name in full_width_fields:
            continue

        field = getField(column)
        setattr(ProfileForm, field_name, field)

    for field_name in full_width_fields:
        if field_name in Users.__table__.columns:
            column = Users.__table__.columns[field_name]
            field = getField(column)
            setattr(ProfileForm, field_name, field)

    form = ProfileForm(obj=current_user)

    if form.validate_on_submit():
        readonly_fields.append("password")
        excluded_fields = readonly_fields
        for field_name, field_value in form.data.items():
            if field_name not in excluded_fields:
                setattr(current_user, field_name, field_value)

        db.session.commit()
        return redirect(url_for('home_blueprint.profile'))
    
    context = {
        'segment': 'profile',
        'form': form,
        'readonly_fields': readonly_fields,
        'full_width_fields': full_width_fields,
    }
    return render_template('pages/profile.html', **context)

@blueprint.route('/<template>')
@login_required
def route_template(template):

    try:

        if not template.endswith('.html'):
            template += '.html'

        # Detect the current page
        segment = get_segment(request)

        # Serve the file (if exists) from app/templates/home/FILE.html
        return render_template("home/" + template, segment=segment)

    except TemplateNotFound:
        return render_template('home/page-404.html'), 404

    except:
        return render_template('home/page-500.html'), 500

# Helper - Extract current page name from request
def get_segment(request):

    try:

        segment = request.path.split('/')[-1]

        if segment == '':
            segment = 'index'

        return segment

    except:
        return None

# Custom template filter
@blueprint.app_template_filter("replace_value")
def replace_value(value, arg):
    return value.replace(arg, " ").title()