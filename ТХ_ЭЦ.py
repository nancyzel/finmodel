import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ТХ_ЭЦ')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')
df_help2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_лизинг')
app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'char': row.iloc[2],
    'numb': row.iloc[3],
    'mean': row.iloc[4],
    'place': row.iloc[5],

    'quantity': row.iloc[6],
    'eur': row.iloc[7],
    'chin': row.iloc[8],
    'rus': row.iloc[9],
    'country': row.iloc[10],

    'price': row.iloc[11],
    'summa': row.iloc[12],
    'lease': row.iloc[13],
    'prod': row.iloc[14],
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование оборудования',

        'field': 'input-data',
        'editable': True,
    },

    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
        'editable': True,

    },
    {
        'headerName': 'Основная характеристика',

        'field': 'char',
        'editable': True,

    },

    {
        'headerName': '№ по схеме ТХ',
        'field': 'numb',
        'editable': True,

    },
    {
        'headerName': 'Назначение',

        'field': 'mean',
        'editable': True,

    },
    {
        'headerName': 'Место расположения',

        'field': 'place',
        'editable': True,


    },
    {
        'headerName': 'Количество',

        'field': 'quantity',

    },
    {
        'headerName': 'Стоимость, за 1 МВт',
        'children': [
            {
                'field': 'eur', 'headerName': 'Европа, EUR',
                'editable': True,
            },
            {
                'field': 'chin', 'headerName': 'Китай, EUR',
                'editable': True,
            },
            {
                'field': 'rus', 'headerName': 'Россия, RUR',
                'editable': True,
            },
        ]

    },
    {
        'headerName': 'Страна',
        'field': 'country',
        'editable': True,

    },

    {
        'headerName': 'Цена за 1 МВт, руб. с НДС',
        'field': 'price',

    },
    {
        'headerName': 'Сумма, руб. с НДС', 'field': 'summa'
    },
    {
        'headerName': 'Лизинг',
        'field': 'lease',
        'editable': True,
    },
    {
        'headerName': 'Производитель',
        'field': 'prod',
        'editable': True,
    },
]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height": 150, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[
                {
                'a1': 'Производственная мощность в год, тонн',
                'a2': df_help['var2'].get(df_help['key'].get(0)-1)
                },
                {
                    'a1': 'Установленная электрическая мощность, МВт',
                    'a2': df_help1['var'].get(df_help1['key'].get(0) - 1),
                }
            ],
            columnDefs=[
                {
                    'headerName': 'Оборудование ЭЦ',
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
            style={"height":400},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},


            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},


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
    '20k': 20000,
    '40k': 40000,
    '80k': 80000,
    '120k': 120000,
    '240k': 240000,
    '360k': 360000,
}
dict1={
    'eur': 1.3,
    'chin':1.3,
    'rus': 1.1
}
sp=['20k','40k','80k','120k','240k','360k']
cur=[95.00, 85.00, 1]
obor=[1353174,0,0]
u=5

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
    if data[0]['country'] == 'Европа':

        data[0]['price'] = data[0]['eur'] * cur[0]

    elif data[0]['country'] == 'Китай':

        data[0]['price'] = data[0]['chin'] * cur[1]

    elif data[0]['country'] == 'Россия':

        data[indr]['price'] = data[indr]['rus'] * cur[2]
    data[0]['summa'] = data[0]['quantity'] * data[0]['price'] * u
    data[1]['summa'] = data[0]['summa']
    if data[0]['lease'] == 'Лизинг':
        data[2]['summa'] = data[1]['summa']
    else:
        data[2]['summa'] = 0



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
        dff.to_excel(writer, sheet_name="ТХ_ЭЦ", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)