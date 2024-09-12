import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='СМР_ПК')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    'floor': row[2],
    'numb': row[3],
    '20k': row.iloc[4],
    '40k': row.iloc[5],
    '80k': row.iloc[6],
    '120k': row.iloc[7],
    '240k': row.iloc[8],
    '360k': row.iloc[9],
    'output-data': row.iloc[10],
    'price': row.iloc[11],
    'summa': row.iloc[12]} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование здания/сооружения',

        'field': 'input-data',
    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
    },
    {
        'headerName': 'Этажность',
        'field': 'floor',
    },
    {
        'headerName': '№ по схеме ГП',
        'field': 'numb',
    },
    {
        'headerName': 'Производственная мощность в год, тонн',
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

            },
            {
                'field': '120k', 'headerName': '120000',

            },
            {
                'field': '240k', 'headerName': '240000',

            },
            {
                'field': '360k', 'headerName': '360000',

            },
        ]

    },
    {
        'headerName': 'Общая площадь, м2',
        'field': 'output-data',
    },
    {
        'headerName': 'Цена, руб. с НДС',
        'field': 'price',
        'editable': True,
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
                    'headerName': 'Здания и сооружения',
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
            style={"height":800},
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
sp=['20k','40k','80k','120k','240k','360k']

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
    if ind in dict:
        if indr in [0, 2, 5, 10]:
            for i in range(6):
                data[indr][sp[i]] = cell_changed[0]['value']
        elif indr in [1, 3, 4, 7, 8]:
            if ind == '20k':
                data[indr]['40k'] = data[indr]['20k']
                for i in range(2, 6):
                    data[indr][sp[i]] = 1.3 * data[indr][sp[i - 1]]
        else:
            data[indr]['20k'] = data[indr]['40k'] / 40000 * 20000
            for i in range(2, 6):
                data[indr][sp[i]] = data[indr]['40k'] / 40000 * dict[sp[i]]

    for i in range(6):
        data[22][sp[i]]=sum([data[j][sp[i]] for j in range(22)])
    for i in range(22):
        data[i]['output-data']=data[i][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[i]['summa']=1.3*float(data[i]['output-data'])*float(data[i]['price'])
    data[22]['summa']=sum([data[j]['summa'] for j in range(22)])

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
        dff.to_excel(writer, sheet_name="СМР_ПК", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)