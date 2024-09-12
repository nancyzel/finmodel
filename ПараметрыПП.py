import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ПараметрыПП')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='ТХ_ПК')
df_help2=pd.read_excel('test.xlsx', sheet_name='Параметры_макроэкпарам')
df_help3=pd.read_excel('test.xlsx', sheet_name='СлужСпр_рабочдавл')
df_help4=pd.read_excel('test.xlsx', sheet_name='СлужСпр_колвоферм')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'form': row.iloc[1],
    'measure': row[2],
    'znach1': row[3],
    'znach2': row.iloc[4],
    'znach3': row.iloc[5],
    'znach4': row.iloc[6],
    'znach5': row.iloc[7],
    'znach6': row.iloc[8],
    'znach7': row.iloc[9],
    'znach8': row.iloc[10],
    'znach9': row.iloc[11],
    'znach10': row.iloc[12],
    'znach11': row.iloc[13],

} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': '',

        'field': 'input-data',
    },
    {
        'headerName': '',
        'field': 'form',

    },
    {
        'headerName': '',
        'field': 'measure',

    },
    {
        'headerName': '',
        'field': 'znach1',

    },
    {
        'headerName': '',
        'field': 'znach2',

    },
    {
        'headerName': '',
        'field': 'znach3',

    },
    {
        'headerName': '',
        'field': 'znach4',

    },
    {
        'headerName': '',
        'field': 'znach5',

    },
    {
        'headerName': '',
        'field': 'znach6',

    },
    {
        'headerName': '',
        'field': 'znach7',

    },
    {
        'headerName': '',
        'field': 'znach8',

    },
    {
        'headerName': '',
        'field': 'znach9',

    },
    {
        'headerName': '',
        'field': 'znach10',

    },
    {
        'headerName': '',
        'field': 'znach11',

    },
]



app.layout = html.Div(


    [
        html.Div(id='try'),
        dag.AgGrid(
            style={"height": 100, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[{
                'a1': 'Производственная мощность в год, тонн',
                'a2': df_help['var2'].get(df_help['key'].get(0)-1),
            }],
            columnDefs=[
                {
                    'headerName': 'Параметры производственного процесса',
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
sp1=['znach1', 'znach2', 'znach3','znach4','znach5','znach6','znach7','znach8','znach9','znach10','znach11']
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


@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=False,
)

def update_row_data(cell_changed,data):
    data[3]['znach7']=data[9]['znach7']/data[9]['znach4']
    for el in sp1:
        data[2][el]=df_help1['char'].get(9)
        data[4][el]=data[1][el]*data[2][el]*data[3][el]
        data[5][el]=df_help1['numb'].get(9)
        data[6][el]=df_help2['value'].get(5)
        data[7][el]=df_help2['value'].get(3)
        data[8][el]=data[4][el]*data[5][el]*data[6][el]*data[7][el]/1000
    data[12]['measure']=df_help3['key'].get(0)
    data[12]['form']=data[12]['measure']
    data[13]['form']=data[5]['znach1']
    if int(df_help4['var'].get(df_help4['key'].get(0)-1))+int(data[13]['form'])<=0:
        data[13]['measure']=1
    else:
        data[13]['measure']=df_help4['var'].get(df_help4['key'].get(0)-1)+data[13]['form']
    data[14]['form']=data[5]['znach1']*data[2]['znach1']
    data[14]['measure']=data[13]['measure']*data[2]['znach1']
    data[15]['form']=data[3][[key for key, value in dict1.items() if value == data[12]['form']][0]]
    data[15]['measure']=data[15]['znach3']/100*data[15]['form']
    data[16]['form']=data[1][[key for key, value in dict1.items() if value == data[12]['form']][0]]
    data[16]['measure']=data[16]['znach3']/100*data[15]['measure']*data[1][[key for key, value in dict1.items() if value == data[12]['measure']][0]]*data[3][[key for key, value in dict1.items() if value == data[12]['measure']][0]]
    data[17]['form']=data[14]['form']*data[16]['form']*data[6]['znach1']*data[7]['znach1']*data[15]['form']/1000
    data[17]['measure']=data[14]['measure']*data[16]['measure']*data[6]['znach1']*data[7]['znach1']*data[15]['measure']*data[3][[key for key, value in dict1.items() if value == data[12]['measure']][0]]/1000

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
        dff.to_excel(writer, sheet_name="ПараметрыПП", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)