import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_0')
df1=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_1')
df2=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_2')
df3=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_3')
df4=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_4')
df5=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_5')
df6=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_6')
df7=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_7')
df8=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_8')
df9=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_9')
df10=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_10')
df11=pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_11')

df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='ТХ_ПК')
df_help2=pd.read_excel('test.xlsx', sheet_name='Параметры_макроэкпарам')
df_help3=pd.read_excel('test.xlsx', sheet_name='СлужСпр_рабочдавл')
df_help4=pd.read_excel('test.xlsx', sheet_name='СлужСпр_колвоферм')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'descript': row.iloc[1],
    'measure': row.iloc[2],
    '1': row.iloc[3],
    '2': row.iloc[4],
    '3': row.iloc[5],
    '4': row.iloc[6],
    '5': row.iloc[7],
    '6': row.iloc[8],
    '7': row.iloc[9],
    '8': row.iloc[10],
    '9': row.iloc[11],
    '10': row.iloc[12],
    '11': row.iloc[13],

} for ind, row in df.iterrows()]
columnDefs=[
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',


    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]

data1=[{
    'concat': row.iloc[0],
    'input-data':row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df1.iterrows()]
columnDefs1=[
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',


    },
    {
        'headerName': '',
        'field': 'measure',


    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]

data2 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df2.iterrows()]
columnDefs2 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]

data3 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df3.iterrows()]
columnDefs3 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',
    },
]
data4 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df4.iterrows()]
columnDefs4 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',

    },
]
data5 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df5.iterrows()]
columnDefs5 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',

    },
]
data6 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df6.iterrows()]
columnDefs6 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]
data7 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df7.iterrows()]
columnDefs7 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]
data8 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df8.iterrows()]
columnDefs8 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',

    },
]
data9 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df9.iterrows()]
columnDefs9 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '2',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '3',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '4',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '5',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '6',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '7',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '8',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '9',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '10',
        'editable': True,

    },
    {
        'headerName': '',
        'field': '11',
        'editable': True,

    },
]
data10 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df10.iterrows()]
columnDefs10 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',

    },
]
data11 = [{
    'concat': row.iloc[0],
    'input-data': row.iloc[1],
    'descript': row.iloc[2],
    'measure': row.iloc[3],
    '1': row.iloc[4],
    '2': row.iloc[5],
    '3': row.iloc[6],
    '4': row.iloc[7],
    '5': row.iloc[8],
    '6': row.iloc[9],
    '7': row.iloc[10],
    '8': row.iloc[11],
    '9': row.iloc[12],
    '10': row.iloc[13],
    '11': row.iloc[14],

} for ind, row in df11.iterrows()]
columnDefs11 = [
    {
        'headerName': '',

        'field': 'concat',

    },
    {
        'headerName': '',

        'field': 'input-data',

    },
    {
        'headerName': '',
        'field': 'descript',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': '1',

    },
    {
        'headerName': '',
        'field': '2',

    },
    {
        'headerName': '',
        'field': '3',

    },
    {
        'headerName': '',
        'field': '4',

    },
    {
        'headerName': '',
        'field': '5',

    },
    {
        'headerName': '',
        'field': '6',

    },
    {
        'headerName': '',
        'field': '7',

    },
    {
        'headerName': '',
        'field': '8',

    },
    {
        'headerName': '',
        'field': '9',

    },
    {
        'headerName': '',
        'field': '10',

    },
    {
        'headerName': '',
        'field': '11',

    },
]

app.layout = html.Div(


    [
        html.Div(id='try'),
        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': 'Расчет удельных энергозатрат на выращивание биомассы в производственных ферментерах',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height":500},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},


            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },


        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table1',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '1. Полезная вводимая энергия с насосом',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table1',
            rowData=data1,
            columnDefs=columnDefs1,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table2',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '2. Полезная энергия на компримирование воздуха',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table2',
            rowData=data2,
            columnDefs=columnDefs2,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table3',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '3. Суммарная полезная энергия',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table3',
            rowData=data3,
            columnDefs=columnDefs3,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table4',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '4. Концентрация биомассы в ферментёре',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table4',
            rowData=data4,
            columnDefs=columnDefs4,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table5',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '5. Удельная вводимая энергия',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table5',
            rowData=data5,
            columnDefs=columnDefs5,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table6',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '6. Объёмный коэффициент массопередачи',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table6',
            rowData=data6,
            columnDefs=columnDefs6,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table7',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '7. Движущая сила процесса абсорбции кислорода',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table7',
            rowData=data7,
            columnDefs=columnDefs7,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table8',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '8. Скорость сорбции кислорода в процессе биосинтеза',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table8',
            rowData=data8,
            columnDefs=columnDefs8,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table9',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '9. Продуктивность процесса',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table9',
            rowData=data9,
            columnDefs=columnDefs9,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table10',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '10. Производительность',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table10',
            rowData=data10,
            columnDefs=columnDefs10,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),

        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table11',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': '11. Удельные энергозатраты',
                    'field': 'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                }
            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height": 500},
            id='computed-table11',
            rowData=data11,
            columnDefs=columnDefs11,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
            },

        ),



        dbc.Col
        (
            [
                dbc.Button(
                    id="save-btn",
                    children="Save Table",
                    color="primary",
                    size="md",
                ),
            ],
            width=3,
        ),
        dbc.Row(
            dbc.Alert(children=None,
                      color="success",
                      id='alerting',
                      is_open=False,
                      duration=2000,
                      style={'width':'18rem'}
            ),
        ),

        html.Div(id="text-field")
    ],
    style={
        'textAlign': 'center',
    },

)

dict={
    '20k': 20000,
    '40k': 40000,
    '80k': 80000,
    '120k': 120000,
    '240k': 240000,
    '360k': 360000,
}
sp=['20k','40k','80k','120k','240k','360k']
sp1 = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11']





@callback(
    #Output('text-field', 'children'),
    Output('computed-table1', 'rowData'),
    Input('computed-table1', 'cellValueChanged'),
    State('computed-table1', 'rowData'),
    State('computed-table', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    for el in sp1:
        data[5][el]=data1[2][el]
        data[0][el]=(data[2][el]*data[3][el]*data[4][el])/(102*3600)*data[5][el]
        data[6][el]=data[0][el]/(data[5][el]*data[6][el])
    for i in range(7):
        try:
            data[i]['concat']=data[i]['input-data']+data[i]['descript']+data[i]['measure']+str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat']=''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table2', 'rowData'),
    Input('computed-table2', 'cellValueChanged'),
    State('computed-table2', 'rowData'),
    State('computed-table', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    for el in sp1:
        data[5][el]=data1[0][el]
        data[2][el]=data[5][el]/data[3][el]
        data[0][el]=(data[2][el]*data[4][el])*0.83*(data[5][el]**0.286-1)
        data[6][el]=data[0][el]/data[4][el]
    for i in range(7):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table3', 'rowData'),
    Input('computed-table3', 'cellValueChanged'),
    State('computed-table3', 'rowData'),
    State('computed-table1', 'rowData'),
    State('computed-table2', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    for el in sp1:
        data[2][el]=data1[0][el]
        data[3][el]=data2[0][el]
        data[4][el]=data1[6][el]+data2[6][el]
        data[0][el]=data[2][el]+data[3][el]

    for i in range(5):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table5', 'rowData'),
    Input('computed-table5', 'cellValueChanged'),
    State('computed-table5', 'rowData'),
    State('computed-table', 'rowData'),
    State('computed-table3', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    for el in sp1:
        data[2][el]=data2[0][el]
        data[3][el]=data1[1][el]
        data[0][el]=data[2][el]/data[3][el]

    for i in range(4):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table6', 'rowData'),
    Input('computed-table6', 'cellValueChanged'),
    State('computed-table6', 'rowData'),
    State('computed-table5', 'rowData'),


    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    for el in sp1:
        data[2][el]=data1[0][el]
        data[0][el]=data[4][el]*data[2][el]**data[3][el]
    for i in range(5):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table7', 'rowData'),
    Input('computed-table7', 'cellValueChanged'),
    State('computed-table7', 'rowData'),
    State('computed-table', 'rowData'),


    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    for el in sp1:
        data[2][el]=data1[0][el]
        data[3][el]=data[2][el]*10/100
        data[0][el]=(data[2][el]*((data[3][el]-data[4][el])*(data[5][el]/data[6][el]))*10**(-3))*1000
    for i in range(2):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    for i in range(3,7):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table8', 'rowData'),
    Input('computed-table8', 'cellValueChanged'),
    State('computed-table8', 'rowData'),
    State('computed-table6', 'rowData'),
    State('computed-table7', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    for el in sp1:
        data[2][el]=data1[0][el]
        data[3][el]=data2[0][el]
        data[0][el]=data[2][el]*data[3][el]*10**(-3)

    for i in range(4):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data


@callback(
    #Output('text-field', 'children'),
    Output('computed-table9', 'rowData'),
    Input('computed-table9', 'cellValueChanged'),
    State('computed-table9', 'rowData'),
    State('computed-table8', 'rowData'),


    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    data[3]['4'] = data1[0]['4']
    data[1]['4']=data[3]['4']/data[4]['4']
    for el in sp1:
        data[3][el]=data1[0][el]
        data[1][el]=data[3][el]/data[4][el]
        if el!='4':
            data[0][el]=data[1][el]/(data[1]['4']/data[0]['4'])

    for i in range(1,5):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table10', 'rowData'),
    Input('computed-table10', 'cellValueChanged'),
    State('computed-table10', 'rowData'),
    State('computed-table', 'rowData'),
    State('computed-table9', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    for el in sp1:
        data[3][el]=data2[1][el]
        data[4][el]=data1[1][el]
        data[1][el]=data[3][el]*data[4][el]
        data[0][el]=data2[0][el]*data[4][el]

    for i in range(1,5):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table11', 'rowData'),
    Input('computed-table11', 'cellValueChanged'),
    State('computed-table11', 'rowData'),
    State('computed-table3', 'rowData'),
    State('computed-table10', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    for el in sp1:
        data[2][el]=data1[0][el]
        data[3][el]=data2[1][el]
        data[0][el]=data[2][el]/data[3][el]
        data[6][el]=data1[4][el]/data[3][el]

    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table4', 'rowData'),
    Input('computed-table4', 'cellValueChanged'),
    State('computed-table4', 'rowData'),
    State('computed-table10', 'rowData'),
    State('computed-table5', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1, data2):
    data[3]['4'] = data1[1]['4']
    data[4]['4'] = data2[3]['4'] / data[5]['4']
    data[1]['4'] = data[3]['4'] / data[4]['4']
    for el in sp1:
        data[3][el]=data1[1][el]
        data[4][el]=data2[3][el]/data[5][el]
        data[1][el]=data[3][el]/data[4][el]
        if el!='4':
            data[0][el]=data[1][el]*data[0]['4']/data[1]['4']

    for i in range(1,5):
        try:
            data[i]['concat'] = data[i]['input-data'] + data[i]['descript'] + data[i]['measure'] + str(round(data[i]['4'], ndigits=2))
        except:
            data[i]['concat'] = ''
    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    State('computed-table11', 'rowData'),

    prevent_initial_call=False,
)

def update_row_data(cell_changed,data, data1):
    for el in sp1:
        data[7][el]=data1[6][el]


    return data

@app.callback(
    Output("alerting", "is_open"),
    Output("alerting", "children"),
    Output("alerting", "color"),
    Input("save-btn", "n_clicks"),

    State("computed-table", "rowData"),
    State("computed-table1", "rowData"),
    State("computed-table2", "rowData"),
    State("computed-table3", "rowData"),
    State("computed-table4", "rowData"),
    State("computed-table5", "rowData"),
    State("computed-table6", "rowData"),
    State("computed-table7", "rowData"),
    State("computed-table8", "rowData"),
    State("computed-table9", "rowData"),
    State("computed-table10", "rowData"),
    State("computed-table11", "rowData"),

    prevent_initial_call=True,
)


def update_portfolio_stats(n, data,data1,data2,data3,data4,data5,data6,data7,data8,data9,data10,data11):
    dff = pd.DataFrame(data)
    dff1 = pd.DataFrame(data)
    dff2 = pd.DataFrame(data)
    dff3 = pd.DataFrame(data)
    dff4 = pd.DataFrame(data)
    dff5 = pd.DataFrame(data)
    dff6 = pd.DataFrame(data)
    dff7 = pd.DataFrame(data)
    dff8 = pd.DataFrame(data)
    dff9 = pd.DataFrame(data)
    dff10 = pd.DataFrame(data)
    dff11 = pd.DataFrame(data)


    with pd.ExcelWriter('test.xlsx', mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
        dff.to_excel(writer, sheet_name="ЗатрТехЭЭ_0", index=False)
        dff1.to_excel(writer, sheet_name="ЗатрТехЭЭ_1", index=False)
        dff2.to_excel(writer, sheet_name="ЗатрТехЭЭ_2", index=False)
        dff3.to_excel(writer, sheet_name="ЗатрТехЭЭ_3", index=False)
        dff4.to_excel(writer, sheet_name="ЗатрТехЭЭ_4", index=False)
        dff5.to_excel(writer, sheet_name="ЗатрТехЭЭ_5", index=False)
        dff6.to_excel(writer, sheet_name="ЗатрТехЭЭ_6", index=False)
        dff7.to_excel(writer, sheet_name="ЗатрТехЭЭ_7", index=False)
        dff8.to_excel(writer, sheet_name="ЗатрТехЭЭ_8", index=False)
        dff9.to_excel(writer, sheet_name="ЗатрТехЭЭ_9", index=False)
        dff10.to_excel(writer, sheet_name="ЗатрТехЭЭ_10", index=False)
        dff11.to_excel(writer, sheet_name="ЗатрТехЭЭ_11", index=False)

    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)