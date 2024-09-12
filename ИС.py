import dash
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ИС')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'char': row.iloc[2],
    'numb': row.iloc[3],
    'mean': row.iloc[4],
    '20k': row.iloc[5],
    '40k': row.iloc[6],
    '80k': row.iloc[7],
    '120k': row.iloc[8],
    '240k': row.iloc[9],
    '360k': row.iloc[10],
    'long': row.iloc[11],
    '20k1': row.iloc[12],
    '40k1': row.iloc[13],
    '80k1': row.iloc[14],
    '120k1': row.iloc[15],
    '240k1': row.iloc[16],
    '360k1': row.iloc[17],
    'sech': row.iloc[18],
    '20k2': row.iloc[19],
    '40k2': row.iloc[20],
    '80k2': row.iloc[21],
    '120k2': row.iloc[22],
    '240k2': row.iloc[23],
    '360k2': row.iloc[24],
    'price': row.iloc[25],
    'mont': row.iloc[26],
    'summa': row.iloc[27],
} for ind, row in df.iterrows()]
columnDefs=[
    {
        'headerName': 'Наименование инженерной системы',
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
        'headerName': 'Производственная мощность, тыс. тонн/год',
        'children':[
            {
                'field': '20k', 'headerName': '20000',
                'editable':True,
            },
            {
                'field': '40k', 'headerName': '40000',
                'editable': True,
            },
            {
                'field': '80k', 'headerName': '80000',
                'editable': True,
            },
            {
                'field': '120k', 'headerName': '120000',
                'editable': True,
            },
            {
                'field': '240k', 'headerName': '240000',
                'editable': True,
            },
            {
                'field': '360k', 'headerName': '360000',
                'editable': True,
            },
        ]

    },
    {
        'headerName': 'Протяжённость',

        'field': 'long',

    },

    {
        'headerName': 'Производственная мощность, тыс. тонн/год',
        'children': [
            {
                'field': '20k1', 'headerName': '20000',
                'editable': True,
            },
            {
                'field': '40k1', 'headerName': '40000',
                'editable': True,
            },
            {
                'field': '80k1', 'headerName': '80000',
                'editable': True,
            },
            {
                'field': '120k1', 'headerName': '120000',
                'editable': True,
            },
            {
                'field': '240k1', 'headerName': '240000',
                'editable': True,
            },
            {
                'field': '360k1', 'headerName': '360000',
                'editable': True,
            },
        ]

    },
    {
        'headerName': 'Сечение',

        'field': 'sech',

    },
    {
        'headerName': 'Производственная мощность, тыс. тонн/год',
        'children': [
            {
                'field': '20k2', 'headerName': '20000',
                'editable': True,
            },
            {
                'field': '40k2', 'headerName': '40000',
                'editable': True,
            },
            {
                'field': '80k2', 'headerName': '80000',
                'editable': True,
            },
            {
                'field': '120k2', 'headerName': '120000',
                'editable': True,
            },
            {
                'field': '240k2', 'headerName': '240000',
                'editable': True,
            },
            {
                'field': '360k2', 'headerName': '360000',
                'editable': True,
            },
        ]
    },
    {
        'headerName': 'Цена',

        'field': 'price',

    },
    {
        'headerName': 'Монтаж ИС',

        'field': 'mont',

    },
    {
        'headerName': 'Сумма, руб. с НДС', 'field': 'summa'
    },

]



app.layout = html.Div(
    [
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
                    'headerName': 'Инженерные системы',
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
            style={"height":1000},
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
        )


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
    '20k1': 20000,
    '40k1': 40000,
    '80k1': 80000,
    '120k1': 120000,
    '240k1': 240000,
    '360k1': 360000,
}
dict2={
    '20k2': 20000,
    '40k2': 40000,
    '80k2': 80000,
    '120k2': 120000,
    '240k2': 240000,
    '360k2': 360000,
}
sp1=['20k1',
    '40k1',
    '80k1',
    '120k1',
    '240k1',
    '360k1',
]
@callback(
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):

    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']

    if ind in dict:
        data[indr]['long'] = data[indr][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[indr]['summa']=data[indr]['long']*data[indr]['price']*(1+data[indr]['mont'])
    elif ind in dict2:
        data[indr]['price'] = data[indr][[key for key, value in dict2.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[indr]['summa']=data[indr]['long']*data[indr]['price']*(1+data[indr]['mont'])
    elif ind in dict1:
        if indr<12:
            for i in range(3,6):
                data[indr][sp1[i]]=data[indr][sp1[i-1]]*1.2
        data[indr]['sech'] = data[indr][[key for key, value in dict1.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
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
        dff.to_excel(writer, sheet_name="ИС", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)