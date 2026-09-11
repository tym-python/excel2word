# !/usr/bin/env python3
# -*-coding:utf-8 -*-

"""
# File       : web.py
# Time       ：2026-07-17 9:06
# Author     ：t
# Func     ：
"""
import dash
from dash import Dash, html, dcc, Input, Output, State, callback, no_update, clientside_callback
import dash_bootstrap_components as dbc
import base64
import os
import io
from datetime import datetime

# 创建应用
app = Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])

# 布局
app.layout = dbc.Container([
    # 标题
    dbc.Row([
        dbc.Col([
            html.H1("📄 数据填充工具", className="text-center my-4 text-primary")
        ], width=12)
    ]),

    # 主卡片
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardBody([
                    # 1. 选择数据源 Excel
                    html.Label("📊 选择数据源（Excel）", className="fw-bold mt-2"),
                    dcc.Upload(
                        id='upload-excel',
                        children=html.Div([
                            '拖拽或 ',
                            html.A('点击选择 .xlsx 文件', style={'color': '#007bff', 'cursor': 'pointer'})
                        ]),
                        style={
                            'width': '100%',
                            'height': '60px',
                            'lineHeight': '60px',
                            'borderWidth': '2px',
                            'borderStyle': 'dashed',
                            'borderRadius': '5px',
                            'textAlign': 'center',
                            'marginBottom': '10px',
                            'borderColor': '#dee2e6'
                        },
                        multiple=False,
                        accept='.xlsx'
                    ),
                    html.Div(id='excel-filename', className='text-muted mb-3'),

                    # 2. 选择模板 Word
                    html.Label("📝 选择模板（Word）", className="fw-bold mt-3"),
                    dcc.Upload(
                        id='upload-word',
                        children=html.Div([
                            '拖拽或 ',
                            html.A('点击选择 .docx 文件', style={'color': '#007bff', 'cursor': 'pointer'})
                        ]),
                        style={
                            'width': '100%',
                            'height': '60px',
                            'lineHeight': '60px',
                            'borderWidth': '2px',
                            'borderStyle': 'dashed',
                            'borderRadius': '5px',
                            'textAlign': 'center',
                            'marginBottom': '10px',
                            'borderColor': '#dee2e6'
                        },
                        multiple=False,
                        accept='.docx'
                    ),
                    html.Div(id='word-filename', className='text-muted mb-3'),

                    # 3. 指定输出位置（文件夹选择）
                    html.Label("📁 输出文件夹路径", className="fw-bold mt-3"),
                    dbc.Row([
                        dbc.Col([
                            dbc.Input(
                                id='output-path',
                                type='text',
                                placeholder='例如：C:/output 或 ./output',
                                value='',
                                className='mb-3'
                            ),
                        ], width=9),
                        dbc.Col([
                            dbc.Button(
                                '📂 选择文件夹',
                                id='folder-select-btn',
                                color='secondary',
                                className='mb-3',
                                style={'width': '100%'}
                            ),
                        ], width=3),
                    ]),
                    html.Small("点击按钮选择文件夹，将自动填入文件夹名（可作为相对路径）", className='text-muted'),

                    # 隐藏的文件选择器（用于选择文件夹）
                    # 注意：使用小写 html.input，并设置 webkitdirectory 属性
                    html.input(
                        id='folder-picker',
                        type='file',
                        webkitdirectory=True,  # 开启文件夹选择
                        style={'display': 'none'}
                    ),

                    # 4. 开始运行按钮
                    dbc.Button(
                        '🚀 开始运行',
                        id='run-btn',
                        color='success',
                        className='mt-3 w-100',
                        disabled=True
                    ),

                    # 5. 输出结果区域
                    html.Div(id='output-area', className='mt-4')
                ])
            ], className='shadow-sm')
        ], width=8, className='mx-auto')
    ])
], fluid=True, className='py-4')

# 存储文件信息（Excel 和 Word）
app.layout.children.append(dcc.Store(id='store-excel', data={}))
app.layout.children.append(dcc.Store(id='store-word', data={}))


# ---- 客户端回调：点击“选择文件夹”按钮触发隐藏的文件夹选择器 ----
clientside_callback(
    """
    function(click) {
        if (click) {
            document.getElementById('folder-picker').click();
        }
        return window.dash_clientside.no_update;
    }
    """,
    Output('folder-picker', 'value'),
    Input('folder-select-btn', 'n_clicks'),
    prevent_initial_call=True
)

# ---- 客户端回调：选择文件夹后，提取文件夹名并填入输出路径 ----
clientside_callback(
    """
    function(files) {
        if (files && files.length > 0) {
            // 获取第一个文件的 webkitRelativePath，提取第一级目录名
            const path = files[0].webkitRelativePath;
            const folderName = path.split('/')[0];
            return folderName;
        }
        return '';
    }
    """,
    Output('output-path', 'value'),
    Input('folder-picker', 'files'),
    prevent_initial_call=True
)


# ---- 后端回调：上传 Excel ----
@callback(
    [Output('excel-filename', 'children'),
     Output('store-excel', 'data')],
    Input('upload-excel', 'contents'),
    State('upload-excel', 'filename'),
    prevent_initial_call=True
)
def upload_excel(contents, filename):
    if contents is not None:
        children = f"✅ 已选择：{filename}"
        store_data = {'filename': filename, 'content': contents}
        return children, store_data
    return no_update, {}


# ---- 后端回调：上传 Word ----
@callback(
    [Output('word-filename', 'children'),
     Output('store-word', 'data')],
    Input('upload-word', 'contents'),
    State('upload-word', 'filename'),
    prevent_initial_call=True
)
def upload_word(contents, filename):
    if contents is not None:
        children = f"✅ 已选择：{filename}"
        store_data = {'filename': filename, 'content': contents}
        return children, store_data
    return no_update, {}


# ---- 后端回调：启用/禁用按钮 ----
@callback(
    Output('run-btn', 'disabled'),
    Input('store-excel', 'data'),
    Input('store-word', 'data')
)
def toggle_button(excel_data, word_data):
    if excel_data and word_data:
        return False
    return True


# ---- 后端回调：运行处理 ----
@callback(
    Output('output-area', 'children'),
    Input('run-btn', 'n_clicks'),
    State('store-excel', 'data'),
    State('store-word', 'data'),
    State('output-path', 'value'),
    prevent_initial_call=True
)
def run_process(n_clicks, excel_data, word_data, output_path):
    if n_clicks and excel_data and word_data:
        excel_name = excel_data.get('filename', '未知')
        word_name = word_data.get('filename', '未知')

        # 确定输出文件夹路径
        base_dir = os.getcwd()
        if output_path:
            output_dir = os.path.join(base_dir, output_path)
        else:
            output_dir = base_dir

        # 确保目录存在
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            return dbc.Alert(f"❌ 创建输出目录失败：{str(e)}", color='danger')

        # 生成输出文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"填充结果_{timestamp}.docx"
        output_full_path = os.path.join(output_dir, output_filename)

        # 模拟处理（演示时创建一个文本文件）
        try:
            with open(output_full_path, 'w', encoding='utf-8') as f:
                f.write(f"模拟填充结果\n数据源：{excel_name}\n模板：{word_name}\n时间：{timestamp}\n")
        except Exception as e:
            return dbc.Alert(f"❌ 文件生成失败：{str(e)}", color='danger')

        # 显示成功信息
        return dbc.Alert([
            html.H4("✅ 处理完成！", className="alert-heading"),
            html.Hr(),
            html.P(f"📊 数据源文件：{excel_name}"),
            html.P(f"📝 模板文件：{word_name}"),
            html.P(f"📁 输出文件夹：{output_dir}"),
            html.P(f"📄 生成文件：{output_filename}"),
            html.P(f"📂 完整路径：{output_full_path}"),
        ], color='info')
    return html.Div()

if __name__ == '__main__':
    app.run(debug=True,  port=8051)