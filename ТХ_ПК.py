import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ТХ_ПК')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_льгсубс')
df_help2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_колвоферм')



app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'model': row.iloc[1],
    'prod': row.iloc[2],
    'measure': row.iloc[3],
    'char': row.iloc[4],
    'numb': row.iloc[5],
    'mean': row.iloc[6],
    'place': row.iloc[7],
    'dep': row.iloc[8],
    '20k': row.iloc[9],
    '40k': row.iloc[10],
    '80k': row.iloc[11],
    '120k': row.iloc[12],
    '240k': row.iloc[13],
    '360k': row.iloc[14],
    'quantity': row.iloc[15],
    'eur': row.iloc[16],
    'chin': row.iloc[17],
    'rus': row.iloc[18],
    'country': row.iloc[19],
    'price': row.iloc[20],
    'summa': row.iloc[21],
    'lease': row.iloc[22],
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование оборудования',

        'field': 'input-data',
        'editable': True,
    },
    {
        'headerName': 'Модель',

        'field': 'model',
        'editable': True,

    },
    {
        'headerName': 'Производитель',

        'field': 'prod',
        'editable': True,

    },
    {
        'headerName': 'Ед. Изм.',
        'field': 'measure',
        'editable': True,

    },
    {
        'headerName': 'Основная характеристика',

        'field': 'char',
        'editable': True,

    },

    {
        'headerName': '№ по схеме ГП',
        'field': 'numb',
        'editable': True,

    },
    {
        'headerName': 'Назначение',

        'field': 'mean',
        'editable': True,

    },
    {
        'headerName': 'Место расположения',

        'field': 'place',
        'editable': True,


    },
    {
        'headerName': 'Зависимость',

        'field': 'dep',
        'editable': True,

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
        'headerName': 'Количество',

        'field': 'quantity',

    },
    {
        'headerName': 'Стоимость, за шт.',
        'children': [
            {
                'field': 'eur', 'headerName': 'Европа, EUR',
                'editable': True,
            },
            {
                'field': 'chin', 'headerName': 'Китай, EUR',
                'editable': True,
            },
            {
                'field': 'rus', 'headerName': 'Россия, RUR',
                'editable': True,
            },
        ]

    },
    {
        'headerName': 'Страна',
        'field': 'country',
        'editable': True,

    },

    {
        'headerName': 'Цена, руб. с НДС',
        'field': 'price',

    },
    {
        'headerName': 'Сумма, руб. с НДС', 'field': 'summa'
    },
    {
        'headerName': 'Лизинг',
        'field': 'lease',
        'editable': True,
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
                        'a2': str(df_help['var2'].get(df_help['key'].get(0) - 1))
                    },
                    {
                        'a1': 'Льготы и субсидии',
                        'a2': df_help1['var'].get(df_help1['key'].get(0) - 1),
                    }
                ],
                columnDefs=[
                    {
                        'headerName': 'Оборудование ПК',
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
            style={"height":4800},
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
dict1={
    'eur': 1.3,
    'chin':1.3,
    'rus': 1.1
}
sp=['20k','40k','80k','120k','240k','360k']
cur=[95.00, 85.00, 1]
#sp1=['eur', 'chin', 'rus']

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
    if ind in ['20k','40k','80k','120k','240k','360k']:
        if indr in [3,9,10,11,12,13,15,16,19,20, 41, 43,76,77,79,80,82,85,91,92,96] or 21<indr<39 or 46<indr<61 or 61<indr<75:
            data[indr]['20k']=data[indr]['40k']/2
            for i in range(2,6):
                data[indr][sp[i]]=data[indr]['40k']*dict[sp[i]]/40000
        for el in sp:
            data[17][el]=df_help2['var'].get(df_help2['key'].get(0) - 1) if df_help['var2'].get(df_help['key'].get(0) - 1)==20000 else 0
        for i in range(107):
            data[i]['quantity'] = data[i][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
        data[indr]['summa']=data[indr]['quantity']*data[indr]['price']*dict1[data[indr]['country']]

        data[5]['place']=sum([data[j]['summa'] for j in range(6,8)])
        data[8]['place']=sum([data[j]['summa'] for j in range(9,17)])
        data[18]['place']=sum([data[j]['summa'] for j in range(19,39)])
        data[40]['place']=sum([data[j]['summa'] for j in range(41,57)])
        data[57]['place']=sum([data[j]['summa'] for j in range(58,69)])
        data[69]['place']=sum([data[j]['summa'] for j in range(70,81)])
        data[81]['place']=sum([data[j]['summa'] for j in range(82,90)])
        data[90]['place']=sum([data[j]['summa'] for j in range(91,98)])




    elif ind in ['eur','chin','rus','country']:
        sumimpv=0
        sumimp=0

        sumrusv=0
        sumrus=0
        if data[indr]['country']=='Европа':
            data[indr]['ind']=1
            data[indr]['price']=data[indr]['eur']*cur[0]
            data[indr]['summa'] = data[indr]['quantity'] * data[indr]['price'] * 1.3
            if data[indr]['lease']=='Выкуп':
                sumimpv+=data[indr]['summa']
            else:
                sumimp+=data[indr]['summa']
        elif data[indr]['country']=='Китай':
            data[indr]['ind']=2
            data[indr]['price']=data[indr]['chin']*cur[1]
            data[indr]['summa'] = data[indr]['quantity'] * data[indr]['price'] * 1.3
            if data[indr]['lease']=='Выкуп':
                sumimpv+=data[indr]['summa']
            else:
                sumimp+=data[indr]['summa']

        elif data[indr]['country']=='Россия':
            data[indr]['ind']=3
            data[indr]['price']=data[indr]['rus']*cur[2]
            data[indr]['summa'] = data[indr]['quantity'] * data[indr]['price'] * 1.1
            if data[indr]['lease']=='Выкуп':
                sumrusv+=data[indr]['summa']
            else:
                sumrus+=data[indr]['summa']

        data[5]['place']=sum([data[j]['summa'] for j in range(6,8)])
        data[8]['place']=sum([data[j]['summa'] for j in range(9,17)])
        data[18]['place']=sum([data[j]['summa'] for j in range(19,39)])
        data[40]['place']=sum([data[j]['summa'] for j in range(41,57)])
        data[57]['place']=sum([data[j]['summa'] for j in range(58,69)])
        data[69]['place']=sum([data[j]['summa'] for j in range(70,81)])
        data[81]['place']=sum([data[j]['summa'] for j in range(82,90)])
        data[90]['place']=sum([data[j]['summa'] for j in range(91,98)])
        data[99]['summa']=sumrusv+sumrus
        data[100]['summa']=sumimpv+sumimp
        data[98]['summa']=data[99]['summa']+data[100]['summa']
        data[102]['summa']=sumrusv
        data[103]['summa']=sumimpv
        data[101]['summa']=data[102]['summa']+data[103]['summa']
        data[105]['summa']=sumrus
        data[106]['summa']=sumimp
        data[104]['summa']=data[105]['summa']+data[106]['summa']





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
        dff.to_excel(writer, sheet_name="ТХ_ПК", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)