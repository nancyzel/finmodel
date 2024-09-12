
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc

import pandas as pd

df = pd.read_excel('test.xlsx', sheet_name='ЗатрПрочЭЭ')


df_help = pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1 = pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')
df_help2 = pd.read_excel('test.xlsx', sheet_name='СлужСпр_энергцентр')
df_help3 = pd.read_excel('test.xlsx', sheet_name='ЗатрТехЭЭ_0')
df_help4 = pd.read_excel('test.xlsx', sheet_name='Параметры_макроэкпарам')
df_help5 = pd.read_excel('test.xlsx', sheet_name='НормыРасхЭР')


app = Dash(__name__)

data = [{
    'input-data': row.iloc[0],
    'measure': row.iloc[1],
    '20k': row[2],
    '40k': row[3],
    '80k': row.iloc[4],
    '120k': row.iloc[5],
    '240k': row.iloc[6],
    '360k': row.iloc[7],
    'output': row.iloc[8],
    'dop': row.iloc[9],

} for ind, row in df.iterrows()]

columnDefs = [
    {
        'headerName': '',

        'field': 'input-data',
    },
    {
        'headerName': '',
        'field': 'measure',
    },
    {
        'field': '20k', 'headerName': '',
        'editable': True,
    },
    {
        'field': '40k', 'headerName': '',
        'editable': True,
    },
    {
        'field': '80k', 'headerName': '',
        'editable': True,

    },
    {
        'field': '120k', 'headerName': '',
        'editable': True,

    },
    {
        'field': '240k', 'headerName': '',
        'editable': True,

    },
    {
        'field': '360k', 'headerName': '',
        'editable': True,

    },





    {
        'headerName': '',
        'field': 'output',
    },

    {
        'headerName': '', 'field': 'dop'
    },
]



app.layout = html.Div(

    [
        dag.AgGrid(
            style={"height": 150, "width": "100%"},
            id='small-table',
            dashGridOptions={'suppressNoRowsOverlay': True},
            rowData=[{
                'a1': 'Производительность процесса в год, тонн',
                'a2': df_help['var2'].get(df_help['key'].get(0) - 1),
            },
                {
                    'a1': 'Электрическая мощность по потребителям, МВт',
                    'a2': df_help1['var'].get(df_help1['key'].get(0) - 1),
                },
                {
                    'a1': 'Энергетический центр',
                    'a2': df_help2['var'].get(df_help2['key'].get(0) - 1),
                }
            ],
            columnDefs=[
                {
                    'headerName': 'Расчет энергозатрат на прочие нужды ',
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
            style={"height": 500},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable": False},

            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30},
                'headerHeight': 0,
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
                      style={'width': '18rem'}
                      ),
        ),

        html.Div(id="text-field")
    ],
    style={
        'textAlign': 'center',
    },

)

dict = {
    '20k': 20000,
    '40k': 40000,
    '80k': 80000,
    '120k': 120000,
    '240k': 240000,
    '360k': 360000,
}
sp = ['20k', '40k', '80k', '120k', '240k', '360k']
sp1 = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14']


sp2=['0', '0_1','0_2','0_3','0_4', '1','1_1','1_2','1_3','1_4', '2','2_1','2_2','2_3','2_4', '3','4','5','6','7','8','9','10','11', '12', '13', '14']
sp3=['0_1','0_2','0_3','0_4','1_1','1_2','1_3','1_4','2_1','2_2','2_3','2_4']

@callback(
    # Output('text-field', 'children'),
    Output('computed-table', 'rowData'),
    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    prevent_initial_call=False,
)
def update_row_data(cell_changed,data):
    data[5]['input-data']=data[2]['input-data']
    data[13]['measure'] = data[13]['20k']

    for el in sp:
        data[2][el]=data[0][el]+df_help3['4'].get(7)
        data[3][el]=df_help4['value'].get(3)
        data[4][el]=df_help4['value'].get(4)
        data[5][el]=data[2][el]/data[3][el]/data[4][el]
        data[6][el]=data[1][el]*(1-data[13][el])-data[5][el]
        data[7][el]=data[6][el]*data[3][el]*data[2][el]
    el='output'
    data[0][el]=df_help['var2'].get(df_help['key'].get(0) - 1)
    data[1][el]=df_help1['var'].get(df_help1['key'].get(0) - 1)*1000
    data[2][el] = data[0][el] + df_help3['4'].get(7)
    data[3][el] = df_help4['value'].get(3)
    data[4][el] = df_help4['value'].get(4)
    data[5][el] = data[2][el] / data[3][el] / data[4][el]
    data[6][el] = data[1][el] * (1 - data[13][el]) - data[5][el]
    data[7][el] = data[6][el] * data[3][el] * data[2][el]
    data[8][el]=df_help5['price'].get(2)/1000
    data[10][el]=data[7][el]*data[8][el]
    data[11][el]=sum([data[j][el] for j in range(2,8)])/1000
    data[2]['dop']=data[2][el]/data[0][el]/1000
    data[7]['dop']=data[7][el]/data[0][el]/1000
    if df_help2['var'].get(df_help2['key'].get(0) - 1)=='Да':
        data[10][el]=data[7][el]*data[9][el]
    else:
        data[10][el]=data[7][el]*data[8][el]



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
        dff.to_excel(writer, sheet_name="ЗатрПрочЭЭ", index=False)


    return True, "Data Saved! Well done!", "success"


if __name__ == '__main__':
    app.run(debug=True)