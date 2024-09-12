import dash
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc

#import plotly.express as px
import pandas as pd
#import js2py


df=pd.read_excel('test.xlsx', sheet_name='ПИР')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')

app = Dash(__name__)
#dash.register_page(__name__, path='/data1', name='data1')

data=[
    {
        'input-data':df.iloc[0]['input-data'],
        'measure': df.iloc[0]['measure'],
        '20k': df.iloc[0]['20k'],
        '40k': df.iloc[0]['40k'],
        '80k': df.iloc[0]['80k'],
        '120k': df.iloc[0]['120k'],
        '240k': df.iloc[0]['240k'],
        '360k': df.iloc[0]['360k'],
        'summa': df.iloc[0]['summa'],

    },
    {
        'input-data':df.iloc[1]['input-data'],
        'measure': df.iloc[1]['measure'],
        '20k': df.iloc[1]['20k'],
        '40k': df.iloc[1]['40k'],
        '80k': df.iloc[1]['80k'],
        '120k': df.iloc[1]['120k'],
        '240k': df.iloc[1]['240k'],
        '360k': df.iloc[1]['360k'],
        'summa': df.iloc[1]['summa'],

    },
    {
        'input-data': df.iloc[2]['input-data'],
        'measure': df.iloc[2]['measure'],
        '20k': df.iloc[2]['20k'],
        '40k': df.iloc[2]['40k'],
        '80k': df.iloc[2]['80k'],
        '120k': df.iloc[2]['120k'],
        '240k': df.iloc[2]['240k'],
        '360k': df.iloc[2]['360k'],
        'summa': df.iloc[2]['summa'],

    },
    {
        'input-data':df.iloc[3]['input-data'],
        'measure': df.iloc[3]['measure'],
        '20k': df.iloc[3]['20k'],
        '40k': df.iloc[3]['40k'],
        '80k': df.iloc[3]['80k'],
        '120k': df.iloc[3]['120k'],
        '240k': df.iloc[3]['240k'],
        '360k': df.iloc[3]['360k'],
        'summa': df.iloc[3]['summa'],

    },
    {
        'input-data': df.iloc[4]['input-data'],
        'measure': df.iloc[4]['measure'],
        '20k': df.iloc[4]['20k'],
        '40k': df.iloc[4]['40k'],
        '80k': df.iloc[4]['80k'],
        '120k': df.iloc[4]['120k'],
        '240k': df.iloc[4]['240k'],
        '360k': df.iloc[4]['360k'],
        'summa': df.iloc[4]['summa'],

    }

]
columnDefs=[
    {
        'headerName': 'Земля',

        'field': 'input-data',
    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',


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
        'headerName': 'Сумма, руб.', 'field': 'summa'
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
                    'headerName': 'ПИР',
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

@callback(
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=True,
)

def update_row_data(cell_changed, data):

    ind=cell_changed[0]['colId']
    indr=cell_changed[0]['rowIndex']
    data[indr+1][ind] = data[indr][ind] * dict[ind]
    data[4][ind] = data[1][ind]+data[3][ind]
    data[indr+1]['summa']=data[indr+1][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
    data[4]['summa'] = data[1]['summa'] + data[3]['summa']
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
        dff.to_excel(writer, sheet_name="ПИР", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)
