
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='СМР_ЭЦ')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')


app = Dash(__name__)

data=[{
    'input-data':row[0],
    'measure': row[1],
    'floor': row[2],
    'numb': row[3],
    '11m': row.iloc[4],
    '30m': row.iloc[5],
    '60m': row.iloc[6],
    '90m': row.iloc[7],
    '180m': row.iloc[8],
    '270m': row.iloc[9],
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
        'headerName': 'Установленная электрическая мощность, МВт',
        'children':[
            {
                'field': '11m', 'headerName': '11',

            },
            {
                'field': '30m', 'headerName': '30',
                'editable':True,
            },
            {
                'field': '60m', 'headerName': '60',

            },
            {
                'field': '90m', 'headerName': '90',

            },
            {
                'field': '180m', 'headerName': '180',

            },
            {
                'field': '270m', 'headerName': '270',

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
                    'headerName': 'Здания и сооружения ЭЦ',
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
            style={"height":200},
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

    if ind=='30m':
        data[0]['11m']=data[0]['30m']/30*11
        for i in range(2,6):
            data[0][sp[i]]=data[0]['30m'] / 30 * dict[sp[i]]
    for i in range(6):
        data[1][sp[i]]=sum([data[j][sp[i]] for j in range(1)])
    data[0]['output-data']=data[0][[key for key, value in dict.items() if value==df_help1['var'].get(df_help['key'].get(0)-1)][0]]
    data[0]['summa']=float(data[0]['output-data'])*float(data[0]['price'])
    data[1]['summa']=sum([data[j]['summa'] for j in range(1)])

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
        dff.to_excel(writer, sheet_name="СМР_ЭЦ", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)