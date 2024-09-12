import dash
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='СтавкиНалоги')


app = Dash(__name__)

data=[{
    'name':row.iloc[0],
    'stavkacom': row.iloc[1],
    'stavka': row.iloc[2],
    'time': row.iloc[3],
    'comment': row.iloc[4]
} for ind, row in df.iterrows()]
columnDefs=[
    {
        'headerName': 'Наименование',
        'field': 'name',
        'editable': True,
    },
    {
        'headerName': 'Общая ставка',
        'field': 'stavkacom',
        'editable': True,
    },
    {
        'headerName': 'Ставка в ОЭЗ Тольятти, Самарская область, %', 'field': 'stavka',
        'editable': True,
    },
    {
        'headerName': 'Срок действия льготы, лет', 'field': 'time',
        'editable': True,
    },
    {
        'headerName': 'Комментарий', 'field': 'comment',
        'editable': True,
    },

]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height": 50, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[],
            columnDefs=[
                {
                    'headerName': 'Ставки налогов',
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
            style={"height":700, 'width':1200},
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
        dff.to_excel(writer, sheet_name="СтавкиНалоги", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)