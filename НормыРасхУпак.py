
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='НормыРасхУпак')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')

app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'weight': row.iloc[1],
    'measure': row.iloc[2],
    'spend': row.iloc[3],
    'price': row.iloc[4],
    'summa': row.iloc[5]
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование ресурса',

        'field': 'input-data',
        'editable': True,
    },
    {
        'headerName': 'Грузоподъёмность, кг',

        'field': 'weight',
        'editable': True,
    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
        'editable': True,

    },
    {
        'headerName': 'Расход, ед. изм./т.г.п.',
        'field': 'spend',
        'editable': True,
    },
    {
        'headerName': 'Цена за ед. изм., руб. с НДС',
        'field': 'price',
        'editable': True,
    },

    {
        'headerName': 'Сумма за ед. изм., руб. с НДС', 'field': 'summa'
    },

]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height":150, "width":"100%"},
            id='small-table',
            dashGridOptions = {'suppressNoRowsOverlay':True},
            rowData=[
                {
                    'a1':'Производительность процесса в год, тонн',
                    'a2': df_help['var2'].get(df_help['key'].get(0)-1)
                }
            ],
            columnDefs=[
                {
                    'headerName': 'Удельный расход упаковочных материалов',
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
            style={"height":700, "width": 1200},
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

    indr=cell_changed[0]['rowIndex']
    data[0]['spend']=1000/data[0]['weight']
    data[1]['spend']=1000/data[1]['weight']
    data[4]['spend']=data[0]['spend']+data[1]['spend']
    data[indr]['summa']=data[indr]['spend']*data[indr]['price']
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
        dff.to_excel(writer, sheet_name="НормыРасхУпак", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)