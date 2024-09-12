
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='НормыРасхМТР')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'spend_lose': row.iloc[2],
    'price': row.iloc[3],
    'summa': row.iloc[4],
    'spend': row.iloc[5],
    'koef': row.iloc[6]
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
        'headerName': 'Расход с учетом потерь, ед. изм./т.г.п.',
        'field': 'spend_lose',
    },
    {
        'headerName': 'Цена за ед. изм., руб. с НДС',
        'field': 'price',
        'editable': True,
    },

    {
        'headerName': 'Сумма за тонну, руб. с НДС', 'field': 'summa'
    },
    {
        'headerName': 'Расход, ед. изм./т.г.п.',
        'field': 'spend',
        'editable': True,
    },
    {
        'headerName': 'Коэффициент потерь, %',
        'field': 'koef',
    },
]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height":50, "width":"100%"},
            id='small-table',
            dashGridOptions = {'suppressNoRowsOverlay':True},
            rowData=[
            ],
            columnDefs=[
                {
                    'headerName': 'Удельный расход МТР',
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
    State('small-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data, data1):

    indr=cell_changed[0]['rowIndex']
    data[indr]['spend_lose']=data[indr]['spend']*(1+data[indr]['koef'])
    data[indr]['summa']=data[indr]['spend_lose']*data[indr]['price']
    data[20]['summa']=sum([data[j]['summa'] for j in range(20)])

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
        dff.to_excel(writer, sheet_name="НормыРасхМТР", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)