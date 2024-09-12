import dash
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback, dcc
import dash_ag_grid as dag
import dash_bootstrap_components as dbc
import json

#import plotly.express as px
import pandas as pd
#import js2py

#js_add='''function isCellEditable('''
df=pd.read_excel('test.xlsx', sheet_name='Параметры_осннастройки')
df1=pd.read_excel('test.xlsx', sheet_name='Параметры_макроэкпарам')
df2=pd.read_excel('test.xlsx', sheet_name='Параметры_капвлож')
df3=pd.read_excel('test.xlsx', sheet_name='Параметры_реал')
df4=pd.read_excel('test.xlsx', sheet_name='Параметры_произв')
df5=pd.read_excel('test.xlsx', sheet_name='Параметры_расх')
df6=pd.read_excel('test.xlsx', sheet_name='Параметры_финанс')
df7=pd.read_excel('test.xlsx', sheet_name='Параметры_налоги')

df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_доппрод')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_льгсубс')
df_help2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help3=pd.read_excel('test.xlsx', sheet_name='ЗУ_1')
df_help4=pd.read_excel('test.xlsx', sheet_name='СМР_Подг')
df_help5=pd.read_excel('test.xlsx', sheet_name='ПИР')
df_help6=pd.read_excel('test.xlsx', sheet_name='АСУТП')
df_help7=pd.read_excel('test.xlsx', sheet_name='СлужСпр_рабочдавл')
df_help8=pd.read_excel('test.xlsx', sheet_name='ЕСН')
df_help9=pd.read_excel('test.xlsx', sheet_name='СтавкиНалоги')



#print(df.to_string())

app = Dash(__name__)

data=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': str(row.iloc[2]),
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df.iterrows()]
columnDefs=[
    {
        'headerName': 'Основные настройки проекта',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True,
                "valueFormatter": {"function": """d3.format("(.2f")(params.value)"""},
                "cellDataType": 'text',

                "cellEditor": {"function": "NumberInput"},
                "cellEditorParams": {"placeholder": "Enter a number"}
            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True
            },
        ]

    },

]

data1=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df1.iterrows()]
columnDefs1=[
    {
        'headerName': 'Макроэкономические параметры',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True
            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True
            },
        ]

    },

]

data2=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df2.iterrows()]
columnDefs2=[
    {
        'headerName': 'Параметры капитальных вложений',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True
            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True
            },
        ]

    },

]

data3=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df3.iterrows()]
columnDefs3=[
    {
        'headerName': 'Параметры реализации',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True

            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True

            },
        ]

    },

]

data4=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df4.iterrows()]
columnDefs4=[
    {
        'headerName': 'Параметры производства',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True

            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True

            },
        ]

    },

]

data5=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df5.iterrows()]
columnDefs5=[
    {
        'headerName': 'Параметры расходов',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True

            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True

            },
        ]

    },

]

data6=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df6.iterrows()]
columnDefs6=[
    {
        'headerName': 'Параметры финансирования',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True

            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True

            },
        ]

    },

]

data7=[{
    'input':row.iloc[0],
    'measure': row.iloc[1],
    'value': row.iloc[2],
    'choose': row.iloc[3],
    'comment': row.iloc[4]

} for a, row in df7.iterrows()]
columnDefs7=[
    {
        'headerName': 'Налоги',
        'children': [
            {
                'headerName': 'Показатель',
                'field': 'input',
            },
            {
                'headerName': 'Ед. изм.', 'field': 'measure',
            },
            {
                'headerName': 'Значение', 'field': 'value',
            },
            {
                'headerName': 'Выбор значения', 'field': 'choose',
                'editable': True

            },
            {
                'headerName': 'Комментарии, пояснения, предложения', 'field': 'comment',
                'editable': True

            },
        ]

    },
]

app.layout = html.Div(
    [


        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': 'Основные параметры проекта',
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
            style={"height":400, "width": '100%'},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},
        ),
        dag.AgGrid(
            style={"height":400, "width": '100%'},
            id='computed-table1',
            rowData=data1,
            columnDefs=columnDefs1,
            defaultColDef={"sortable":False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},


        ),
        dag.AgGrid(
            style={"height": 700, "width": '100%'},
            id='computed-table2',
            rowData=data2,
            columnDefs=columnDefs2,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 500, "width": '100%'},
            id='computed-table3',
            rowData=data3,
            columnDefs=columnDefs3,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 400, "width": '100%'},
            id='computed-table4',
            rowData=data4,
            columnDefs=columnDefs4,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 550, "width": '100%'},
            id='computed-table5',
            rowData=data5,
            columnDefs=columnDefs5,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 500, "width": '100%'},
            id='computed-table6',
            rowData=data6,
            columnDefs=columnDefs6,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 800, "width": '100%'},
            id='computed-table7',
            rowData=data7,
            columnDefs=columnDefs7,
            defaultColDef={"sortable": False},
            columnSize="sizeToFit",

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),


        dbc.Col
        (
            [
                dbc.Button(

                    id="save-btn",
                    children="Save Tables",
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
        html.Div(id='text-field'),



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
sp=[20000,40000,80000,120000,240000,360000]


@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']
    return ind


'''
    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    data[4]['value']=df_help['var'].get(df_help['key'].get(0)-1)
    data[5]['value']=df_help1['var'].get(df_help1['key'].get(0)-1)

    return data
'''
@callback(
    #Output('text-field', 'children'),
    Output('computed-table1', 'rowData'),
    Input('computed-table1', 'cellValueChanged'),
    State('computed-table1', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    return data


@callback(
    #Output('text-field', 'children'),
    Output('computed-table2', 'rowData'),
    Input('computed-table2', 'cellValueChanged'),
    State('computed-table2', 'rowData'),
    State('computed-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data, data1):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    data[0]['value']=df_help3[[key for key, value in dict.items() if value==data1[0]['value']][0]].get(0) if data1[0]['value'] in sp else 0
    data[1]['value']=df_help4[[key for key, value in dict.items() if value==data1[0]['value']][0]].get(0) if data1[0]['value'] in sp else 0
    data[2]['value']=df_help5[[key for key, value in dict.items() if value==data1[0]['value']][0]].get(0) if data1[0]['value'] in sp else 0
    data[3]['value']=df_help5[[key for key, value in dict.items() if value==data1[0]['value']][0]].get(2) if data1[0]['value'] in sp else 0
    data[4]['value']=df_help6[[key for key, value in dict.items() if value==data1[0]['value']][0]].get(0) if data1[0]['value'] in sp else 0


    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table3', 'rowData'),
    Input('computed-table3', 'cellValueChanged'),
    State('computed-table3', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    return data


@callback(
    #Output('text-field', 'children'),
    Output('computed-table4', 'rowData'),
    Input('computed-table4', 'cellValueChanged'),
    State('computed-table4', 'rowData'),
    State('computed-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data, data1):
    data[0]['value']=df_help7['var'].get(df_help7['key'].get(0)-1)

    return data

@callback(
    #Output('text-field', 'children'),
    Output('computed-table5', 'rowData'),
    Input('computed-table5', 'cellValueChanged'),
    State('computed-table5', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    return data


@callback(
    #Output('text-field', 'children'),
    Output('computed-table6', 'rowData'),
    Input('computed-table6', 'cellValueChanged'),
    State('computed-table6', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    return data


@callback(
    #Output('text-field', 'children'),
    Output('computed-table7', 'rowData'),
    Input('computed-table7', 'cellValueChanged'),
    State('computed-table7', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):
    ind=cell_changed[0]['colId']

    if ind=='choose':
        data[indr]['value']=data[indr]['choose']
    data[1]['value']=df_help8['basa'].get(0)
    data[2]['value']=df_help8['over'].get(0)
    data[3]['value']=df_help8['basa'].get(1)
    data[4]['value']=df_help8['over'].get(1)
    data[5]['value']=df_help8['basa'].get(2)
    if df_help1['key'].get(0) == 1:
        data[6]['value'] = df_help9['stavka'].get(1)
    else:
        data[6]['value'] = df_help9['stavkacom'].get(1)
    if df_help1['key'].get(0) == 1:
        data[7]['value'] = df_help9['stavka'].get(0)
    else:
        data[7]['value'] = df_help9['stavkacom'].get(0)
    for i in range(8, 15):
        if df_help1['key'].get(0)==1:
            data[i]['value']=df_help9['stavka'].get(i-5)
        else:
            data[i]['value']=df_help9['stavkacom'].get(i-5)



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

    prevent_initial_call=True,
)




def update_portfolio_stats(n, data, data1, data2, data3,data4,data5,data6,data7):#data8,data9,data10,data11,data12,data13,data14,data15,data16):
    dff = pd.DataFrame(data)
    dff1 = pd.DataFrame(data1)
    dff2 = pd.DataFrame(data2)
    dff3 = pd.DataFrame(data3)
    dff4 = pd.DataFrame(data4)
    dff5 = pd.DataFrame(data5)
    dff6 = pd.DataFrame(data6)
    dff7 = pd.DataFrame(data7)


    with pd.ExcelWriter('test.xlsx', mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
        dff.to_excel(writer, sheet_name="Параметры_осннастройки", index=False)
        dff1.to_excel(writer, sheet_name="Параметры_макроэкпарам", index=False)
        dff2.to_excel(writer, sheet_name="Параметры_капвлож", index=False)
        dff3.to_excel(writer, sheet_name="Параметры_реал", index=False)
        dff4.to_excel(writer, sheet_name="Параметры_произв", index=False)
        dff5.to_excel(writer, sheet_name="Параметры_расх", index=False)
        dff6.to_excel(writer, sheet_name="Параметры_финанс", index=False)
        dff7.to_excel(writer, sheet_name="Параметры_налоги", index=False)


    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)