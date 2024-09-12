import dash
from dash import Dash, Patch, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='НормыЗапМТР')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='НормыРасхМТР')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'spend': row.iloc[2],
    'zapas10': row.iloc[3],
    'zapasmonth': row.iloc[4],
    'zapasmonthav': row.iloc[5],
    'price': row.iloc[6],
    'summa': row.iloc[7]} for ind, row in df.iterrows()]
columnDefs=[
    {
        'headerName': 'Наименование ресурса',
        'field': 'input-data',

    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
        'editable': True,
    },
    {
        'headerName': 'Часовой расход ресурсов, кг (при объёме ГП т/час)',
        'children':[
            {
                'field': 'spend', 'headerName': df_help['var2'].get(df_help['key'].get(0)-1)/int(data[22]['measure']) / int(data[23]['measure']) / int(data[24]['measure'])
            }
        ]

    },
    {
        'headerName': '10-ти суточный запас МТР',
        'field': 'zapas10',
        'editable': True,
    },
    {
        'headerName': 'Месячный запас МТР',
        'field': 'zapasmonth',
        'editable': True,
    },
    {
        'headerName': 'Среднемесячный запас МТР',
        'field': 'zapasmonthav',

    },
    {
        'headerName': 'Цена за ед. изм.',
        'field': 'price',

    },
    {
        'headerName': 'Сумма, руб./мес.', 'field': 'summa'
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
            }],
            columnDefs=[
                {
                    'headerName': 'Запасы МТР',
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
        #html.Div(id='text-field'),


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
@callback(
    #Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Output('computed-table', 'columnDefs'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    State('computed-table', 'columnDefs'),

    prevent_initial_call=True,
)

def update_row_data(cell_changed, data, column_defs):


    indr = cell_changed[0]['rowIndex']



    if 20<indr<25:
        patched_grid = Patch()
        patched_grid[2]['children'] = column_defs[2]['children'][0]['headerName'] = 20000 / int(data[22]['measure']) / int(data[23]['measure']) / int(data[24]['measure'])
    znach = df_help['var2'].get(df_help['key'].get(0) - 1) / int(data[22]['measure']) / int(data[23]['measure']) / int(data[24]['measure'])
    for i in range(20):
        data[i]['spend']=df_help1['spend_lose'].get([i for i in range(20) if df_help1['input-data'].get(i)==data[i]['input-data']][0])
        data[i]['price']=df_help1['price'].get([i for i in range(20) if df_help1['input-data'].get(i)==data[i]['input-data']][0])

    data[0]['zapas10'] = data[0]['spend'] * 24 * 10
    data[9]['zapas10'] = data[9]['spend'] * 24 * 10
    data[12]['zapas10'] = data[12]['spend'] * 24 * 10
    data[14]['zapas10'] = data[14]['spend'] * 24 * 10
    data[15]['zapas10'] = data[15]['spend'] * 24 * 10
    data[17]['zapas10'] = data[17]['spend'] * 24 * 10
    data[3]['zapasmonth'] = data[3]['spend'] * data[23]['measure']
    data[5]['zapasmonth'] = data[5]['spend'] * data[23]['measure']
    data[8]['zapasmonth'] = data[8]['spend'] * data[23]['measure']
    data[10]['zapasmonth'] = data[10]['spend'] * data[23]['measure']
    data[11]['zapasmonth'] = data[11]['spend'] * data[23]['measure']
    data[13]['zapasmonth'] = data[13]['spend'] * data[23]['measure']
    data[16]['zapasmonth'] = data[16]['spend'] * data[23]['measure']
    data[18]['zapasmonth'] = data[18]['spend'] * data[23]['measure']
    data[19]['zapasmonth'] = data[19]['spend'] * data[23]['measure']
    for i in range(20):
        data[i]['zapasmonthav'] = data[i]['zapas10'] + data[i]['zapasmonth']
        data[i]['summa'] = data[i]['zapasmonthav'] * data[i]['price']
    data[20]['summa'] = sum([data[i]['summa'] for i in range(20)])



    return data, column_defs

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
        dff.to_excel(writer, sheet_name="НормыЗапМТР", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)