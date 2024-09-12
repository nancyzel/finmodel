
from dash import Dash, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc
import requests
import math
import pandas as pd
def get_exchange_rate(base_currency, target_currency):
    url = f"https://open.er-api.com/v6/latest/{base_currency}"
    response = requests.get(url)

    if response.status_code != 200:
        raise Exception("Error: API request unsuccessful")

    info = response.json()
    if target_currency in info['rates']:
        return info['rates'][target_currency]
    else:
        raise Exception(f"Error: Unable to find the rate for {target_currency}")


df=pd.read_excel('test.xlsx', sheet_name='СтоимОборЭЦ')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')

df_help2=pd.read_excel('test.xlsx', sheet_name='СМР_ЭЦ')



base = "USD"
target = "EUR"
rate = get_exchange_rate(base, target)
rate1=get_exchange_rate("EUR", "RUB")

app = Dash(__name__)

data=[{
    'name':row.iloc[0],
    'eur1': row.iloc[1],
    'eur2': row.iloc[2],
    'chin1': row.iloc[3],
    'chin2': row.iloc[4],
    'rus1': row.iloc[5],
    'rus2': row.iloc[6],
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование показателя',

        'field': 'name',
        'editable': True,
    },

    {
        'headerName': 'Европа',
        'children': [
            {
                'field': 'eur1', 'headerName': 'EUR',

            },
            {
                'field': 'eur2', 'headerName': 'RUR',

            }
        ]

    },
    {
        'headerName': 'Китай',
        'children': [
            {
                'field': 'chin1', 'headerName': 'EUR',
                'editable': True,

            },
            {
                'field': 'chin2', 'headerName': 'RUR',
                'editable': True,

            }
        ]

    },
    {
        'headerName': 'Россия',
        'children': [
            {
                'field': 'rus1', 'headerName': 'EUR',
                'editable': True,

            },
            {
                'field': 'rus2', 'headerName': 'RUR',
                'editable': True,

            }
        ]

    }

]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height":250, "width":"100%"},
            id='small-table',
            dashGridOptions = {'suppressNoRowsOverlay':True},
            rowData=[
                {
                    'a1':'Производственная мощность в год, тонн',
                    'a2': df_help['var2'].get(df_help['key'].get(0)-1)
                },
                {
                    'a1': 'Установленная электрическая мощность, МВт',
                    'a2': df_help1['var'].get(df_help['key'].get(0) - 1)
                },
                {
                    'a1': 'Курс евро, руб',
                    'a2': rate1
                },
            ],
            columnDefs=[
                {
                    'headerName': 'Расчёт стоимости капитальных затрат на оборудование ЭЦ',
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
    if ind in ['eur1','eur2','chin1', 'chin2', 'rus1','rus2']:
        data[0]['eur1']=1400*rate
        data[1]['eur1']=math.ceil(0.48*df_help1['var'].get(df_help['key'].get(0) - 1))*data[0]['eur1']*1000
        data[2]['eur1']=df_help2['summa'].get(1)/rate1
        data[3]['eur1']=data[1]['eur1']-data[2]['eur1']
        data[4]['eur1']=data[3]['eur1']/df_help1['var'].get(df_help['key'].get(0) - 1)
        for i in range(5):
            data[i]['eur2']=data[i]['eur1']*rate1
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
        dff.to_excel(writer, sheet_name="СтоимОборЭЦ", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)