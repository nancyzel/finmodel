
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='НормыРасхЭР')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')
df_help2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_рабочдавл')
df_help3=pd.read_excel('test.xlsx', sheet_name='СлужСпр_энергцентр')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'atm1': row.iloc[2],
    'atm2': row.iloc[3],
    'atm3': row.iloc[4],
    'atm4': row.iloc[5],
    'atm5': row.iloc[6],
    'atm6': row.iloc[7],
    'atm7': row.iloc[8],
    'atm8': row.iloc[9],
    'atm9': row.iloc[10],
    'atm10': row.iloc[11],
    'atm11': row.iloc[12],
    'spend': row.iloc[13],
    'price': row.iloc[14],
    'selfprice': row.iloc[15],
    'summa': row.iloc[16]
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование ресурса',

        'field': 'input-data',
        'editable': True,
    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
        'editable': True,

    },

    {
        'headerName': 'Рабочее давление, атм',
        'children':[
            {
                'field': 'atm1', 'headerName': '1',
                'editable': True,

            },
            {
                'field': 'atm2', 'headerName': '2',
                'editable':True,
            },
            {
                'field': 'atm3', 'headerName': '3',
                'editable': True,

            },
            {
                'field': 'atm4', 'headerName': '4',
                'editable': True,

            },
            {
                'field': 'atm5', 'headerName': '5',
                'editable': True,

            },
            {
                'field': 'atm6', 'headerName': '6',
                'editable': True,
            },
            {
                'field': 'atm7', 'headerName': '7',
                'editable': True,
            },
            {
                'field': 'atm8', 'headerName': '8',
                'editable': True,
            },
            {
                'field': 'atm9', 'headerName': '9',
                'editable': True,
            },
            {
                'field': 'atm10', 'headerName': '10',
                'editable': True,
            },
            {
                'field': 'atm11', 'headerName': '11',
                'editable': True,
            }
        ]

    },
    {
        'headerName': 'Расход, ед. изм./т.г.п.',
        'field': 'spend',
    },
    {
        'headerName': 'Цена за ед. изм., руб. с НДС',
        'field': 'price',
        'editable': True,
    },
    {
        'headerName': 'Себестоимость ед. изм.',
        'field': 'selfprice',
    },
    {
        'headerName': 'Сумма за тонну, руб. с НДС', 'field': 'summa'
    },

]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height":250, "width":"100%"},
            id='small-table',
            dashGridOptions = {'suppressNoRowsOverlay':True},
            rowData=[
                {
                    'a1':'Производительность процесса в год, тонн',
                    'a2':str(df_help['var2'].get(df_help['key'].get(0)-1)),
                },
                {
                    'a1':'Установленная электрическая мощность, МВт',
                    'a2':str(df_help1['var'].get(df_help['key'].get(0)-1)),
                },
                {
                    'a1': 'Рабочее давление в ферментерах, атм',
                    'a2': str(df_help2['var'].get(df_help['key'].get(0)-1)),
                },
                {
                    'a1': 'Энергетический центр',
                    'a2': df_help3['var'].get(df_help['key'].get(0)-1),
                }
            ],
            columnDefs=[
                {
                    'headerName': 'Удельный расход энергоресурсов',
                    'field':'a1'
                },
                {
                    'headerName': '',
                    'field': 'a2'
                },

            ],
            columnSize="sizeToFit",

        ),

        dag.AgGrid(
            style={"height":400},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},


            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},

                "animateRows": False
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

        #html.Div(id="text-field")
    ],
    style={
        'textAlign': 'center',
    },

)

dict={
    '11m': 11,
    '30m': 30,
    '60m': 60,
    '90m': 90,
    '180m': 180,
    '270m': 270,
}
sp=['11m','30m','60m','90m','180m','270m']

@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),

    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):

    ind=cell_changed[0]['colId']
    data[0]['spend']=data[0]['atm4']
    data[2]['spend']=data[2]['atm4']
    if df_help3['var'].get(df_help['key'].get(0)-1)=='Да':
        data[1]['selfprice']=0
    else:
        data[1]['selfprice']=data[1]['price']
        data[2]['selfprice']=data[2]['price']
    data[0]['summa']=data[0]['spend']*data[0]['price']
    for i in range(1,5):
        data[i]['summa']=data[i]['spend']*data[i]['selfprice']
    data[5]['summa']=sum([data[j]['summa'] for j in range(5)])

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
        dff.to_excel(writer, sheet_name="НормыРасхЭР", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)