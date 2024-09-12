
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='НормыРасхВода')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='ТХ_ПК')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'measure': row.iloc[1],
    '20k': row.iloc[2],
    '40k': row.iloc[3],
    '80k': row.iloc[4],
    '120k': row.iloc[5],
    '240k': row.iloc[6],
    '360k': row.iloc[7],
    'spend': row.iloc[8],
    'price': row.iloc[9],
    'summa': row.iloc[10],
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
        'headerName': 'Производственная мощность в год, тонн',
        'children': [
            {
                'field': '20k', 'headerName': '20000',
                'editable': True,
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
        'headerName': 'Расход м3/год',
        'field': 'spend',
    },
    {
        'headerName': 'Цена за м3, руб. с НДС',
        'field': 'price',
        'editable': True,
    },

    {
        'headerName': 'Сумма в год, руб. с НДС', 'field': 'summa'
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
                    'a1':'Производственная мощность в год, тонн',
                    'a2': df_help['var2'].get(df_help['key'].get(0)-1)
                }
            ],
            columnDefs=[
                {
                    'headerName': 'Водоснабжение и водоотведение для хозяйственно-бытовых нужд',
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
            style={"height":500},
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

    indr=cell_changed[0]['rowIndex']
    ind=cell_changed[0]['colId']
    if ind in dict:
        ss=0
        for i in range(9, 18):
            ss+=float(df_help1['char'].get(i))*df_help1['20k'].get(i)


        data[0]['20k']=20000/360*0.3
        data[1]['20k']=20000/360*0.3
        data[2]['20k']=20000/360*0.3
        data[3]['20k']=ss
        data[4]['20k']=data[4]['40k']*20000/40000
        for el in ['120k','240k','360k']:
            data[0][el]=data[0]['40k']*dict[el]/40000
            data[1][el]=data[0]['40k']*dict[el]/40000
            data[4][el]=data[4]['40k']*dict[el]/40000
        data[0]['spend']=data[0][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[1]['spend']=data[1][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[4]['spend']=data[4][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[2]['spend']=data[2][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]*300
        data[3]['spend']=data[3][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]*4
    for i in range(1,5):
        data[i]['summa']=data[i]['spend']*data[i]['price']
    data[0]['summa']=sum([data[j]['summa'] for j in range(1,4)])
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
        dff.to_excel(writer, sheet_name="НормыРасхВода", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)