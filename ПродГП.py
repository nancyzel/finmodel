import dash

from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ПродГП')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_доппрод')
df_help2=pd.read_excel('test.xlsx', sheet_name='Параметры_произв')
df_help3=pd.read_excel('test.xlsx', sheet_name='Параметры_реал')
df_help4=pd.read_excel('test.xlsx', sheet_name='СлужСпр_колвоферм')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'summa': row[2],
    '0': row[3],
    '0_1': row.iloc[4],
    '0_2': row.iloc[5],
    '0_3': row.iloc[6],
    '0_4': row.iloc[7],
    '1': row[8],
    '1_1': row.iloc[9],
    '1_2': row.iloc[10],
    '1_3': row.iloc[11],
    '1_4': row.iloc[12],
    '2': row[13],
    '2_1': row.iloc[14],
    '2_2': row.iloc[15],
    '2_3': row.iloc[16],
    '2_4': row.iloc[17],
    '3': row.iloc[18],
    '4': row.iloc[19],
    '5': row.iloc[20],
    '6': row.iloc[21],
    '7': row.iloc[22],
    '8': row.iloc[23],
    '9': row.iloc[24],
    '10': row.iloc[25],
    '11': row.iloc[26],
    '12': row.iloc[27],
    '13': row.iloc[28],
    '14': row.iloc[29],

} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Показатели реализации ГП',

        'field': 'input-data',
    },

    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',

    },
    {
        'headerName': 'ИТОГО',
        'field': 'summa',

    },
    {
        'headerName': '0',
        'field': '0',

    },
    {
        'headerName': 'в том числе',
        'children': [
            {
                'field': '0_1', 'headerName': '1 кв.',

            },
            {
                'field': '0_2', 'headerName': '2 кв.',
            },
            {
                'field': '0_3', 'headerName': '3 кв.',
            },
            {
                'field': '0_4', 'headerName': '4 кв.',
            },
        ]

    },
    {
        'headerName': '1',
        'field': '1',

    },
    {
        'headerName': 'в том числе',
        'children': [
            {
                'field': '1_1', 'headerName': '1 кв.',

            },
            {
                'field': '1_2', 'headerName': '2 кв.',
            },
            {
                'field': '1_3', 'headerName': '3 кв.',
            },
            {
                'field': '1_4', 'headerName': '4 кв.',
            },
        ]

    },
    {
        'headerName': '2',
        'field': '2',

    },
    {
        'headerName': 'в том числе',
        'children': [
            {
                'field': '2_1', 'headerName': '1 кв.',

            },
            {
                'field': '2_2', 'headerName': '2 кв.',
            },
            {
                'field': '2_3', 'headerName': '3 кв.',
            },
            {
                'field': '2_4', 'headerName': '4 кв.',
            },
        ]

    },
    {
        'headerName': '3',
        'field': '3',

    },
    {
        'headerName': '4',
        'field': '4',

    },
    {
        'headerName': '5',
        'field': '5',

    },
    {
        'headerName': '6',
        'field': '6',

    },
    {
        'headerName': '7',
        'field': '7',

    },
    {
        'headerName': '8',
        'field': '8',

    },
    {
        'headerName': '9',
        'field': '9',

    },
    {
        'headerName': '10',
        'field': '10',

    },
    {
        'headerName': '11',
        'field': '11',

    },
    {
        'headerName': '12',
        'field': '12',

    },
    {
        'headerName': '13',
        'field': '13',

    },
    {
        'headerName': '14',
        'field': '14',

    },
]



app.layout = html.Div(


    [
        dag.AgGrid(
            style={"height": 100, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[{
                'a1': 'Производительность процесса в год, тонн',
                'a2': df_help['var2'].get(df_help['key'].get(0)-1),
            },
{
                'a1': 'Реализация дополнительной продукции',
                'a2': df_help1['var'].get(df_help1['key'].get(0)-1),
            }
            ],
            columnDefs=[
                {
                    'headerName': 'Расчёт объёма реализованной ГП',
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
sp1=['0', '1', '2', '3','4','5','6','7','8','9','10','11', '12', '13', '14']
dict1={'znach1':1,
       'znach2':2,
       'znach3':3,
       'znach4':4,
       'znach5':5,
       'znach6':6,
       'znach7':7,
       'znach8':8,
       'znach9':9,
       'znach10':10,
       'znach11':11}
sp2=['0', '0_1','0_2','0_3','0_4', '1','1_1','1_2','1_3','1_4', '2','2_1','2_2','2_3','2_4', '3','4','5','6','7','8','9','10','11', '12', '13', '14']

@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=False,
)

def update_row_data(cell_changed,data):

    for el in sp1[:3]:
        if data[0][el]==0.00:
            data[3][el]=0
        data[4][el]=data[0][el]*data[3][el]
        data[5][el]=data[4][el]*df_help2['value'].get(3)
        data[6][el]=data[4][el]-data[5][el]
        data[9][el]=data[6][el]+data[7][el]-data[8][el]
        data[10][el]=data[9][el]*df_help3['value'].get(6)*0.5
        data[11][el]=data[9][el]-data[10][el]
        data[12][el]=data[8][el]*df_help3['value'].get(7)
        data[13][el]=data[8][el]*(1-df_help3['value'].get(7))
        if not df_help1['var'].get(df_help1['key'].get(0)-1)=='Да':
            data[14][el]=0
        data[15][el]=data[14][el]*df_help3['value'].get(0)
        data[17][el]=data[15][el]-data[16][el]
        try:
            data[1][el]=data[3][el]/(data[3][el]+data[14][el])
        except:
            data[1][el]=0
        try:
            data[2][el]=data[14][el]/(data[3][el]+data[14][el])
        except:
            data[2][el]=0
    el='summa'
    try:
        data[1][el] = data[3][el] / (data[3][el] + data[14][el])
    except:
        data[1][el] = 0
    try:
        data[2][el] = data[14][el] / (data[3][el] + data[14][el])
    except:
        data[2][el] = 0
    for i in range(3,17):
        data[i]['summa']=sum([data[i][x] for x in sp1])

    for el in sp1[3:-1]:
        if data[0][el]==0.00:
            data[3][el]=0
        data[4][el]=data[0][el]*data[3][el]
        data[5][el]=data[4][el]*df_help2['value'].get(3)
        data[6][el]=data[4][el]-data[5][el]
        data[7][el]=data[8][sp1[sp1.index(el)-1]]
        data[8][el]=data[6][el]/12*df_help3['value'].get(8)
        data[9][el]=data[6][el]+data[7][el]-data[8][el]
        data[10][el]=data[9][el]*df_help3['value'].get(6)*0.5
        data[11][el]=data[9][el]-data[10][el]
        data[12][el]=data[8][el]*df_help3['value'].get(7)
        data[13][el]=data[8][el]*(1-df_help3['value'].get(7))
        if not df_help1['var'].get(df_help1['key'].get(0)-1)=='Да':
            data[14][el]=0
        data[15][el]=data[14][el]*df_help3['value'].get(0)
        data[17][el]=data[15][el]-data[16][el]
        try:
            data[1][el]=data[3][el]/(data[3][el]+data[14][el])
        except:
            data[1][el]=0
        try:
            data[2][el]=data[14][el]/(data[3][el]+data[14][el])
        except:
            data[2][el]=0
    el='14'
    if data[0][el] == 0.00:
        data[3][el] = 0
    data[4][el] = data[0][el] * data[3][el]
    data[5][el] = data[4][el] * df_help2['value'].get(3)
    data[6][el] = data[4][el] - data[5][el]
    data[7][el] = data[8][sp1[sp1.index(el) - 1]]
    data[8][el] = data[8][sp1[sp1.index(el) - 1]]
    data[9][el] = data[6][el] + data[7][el] - data[8][el]
    data[10][el] = data[9][el] * df_help3['value'].get(6) * 0.5
    data[11][el] = data[9][el] - data[10][el]
    data[12][el] = data[8][el] * df_help3['value'].get(7)
    data[13][el] = data[8][el] * (1 - df_help3['value'].get(7))
    if not df_help1['var'].get(df_help1['key'].get(0) - 1) == 'Да':
        data[14][el] = 0
    data[15][el] = data[14][el] * df_help3['value'].get(0)
    data[17][el] = data[15][el] - data[16][el]
    try:
        data[1][el] = data[3][el] / (data[3][el] + data[14][el])
    except:
        data[1][el] = 0
    try:
        data[2][el] = data[14][el] / (data[3][el] + data[14][el])
    except:
        data[2][el] = 0
    el='summa'
    try:
        data[1][el] = data[3][el] / (data[3][el] + data[14][el])
    except:
        data[1][el] = 0
    try:
        data[2][el] = data[14][el] / (data[3][el] + data[14][el])
    except:
        data[2][el] = 0
    for i in range(3,17):
        data[i]['summa']=sum([data[i][x] for x in sp2])


    return data



@app.callback(
    Output("alerting", "is_open"),
    Output("alerting", "children"),
    Output("alerting", "color"),
    Input("save-btn", "n_clicks"),

    State("computed-table", "rowData"),

    prevent_initial_call=True,
)

def update_portfolio_stats(n, data):
    dff = pd.DataFrame(data)

    with pd.ExcelWriter('test.xlsx', mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
        dff.to_excel(writer, sheet_name="ПродГП", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)