
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc

#import plotly.express as px
import pandas as pd
#import js2py

#js_add='''function isCellEditable('''
df=pd.read_excel('test.xlsx', sheet_name='ЗУ_1')
df1=pd.read_excel('test.xlsx', sheet_name='ЗУ_2')
#print(df.to_string())
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_льгсубс')
app = Dash(__name__)

data1=[
    {

        'input-data':'Годовая ставка аренды ЗУ',
        'measure': 'руб./м2',
        'output-data': df1.iloc[0]['output-data'],
        'arenda': df1.iloc[0]['arenda'],
        'vikup_price': df1.iloc[0]['vikup_price'],
        'arenda_price': df1.iloc[0]['arenda_price'],
        'vikup': df1.iloc[0]['vikup'],
    },
    {
        'input-data':'Стоимость выкупа ЗУ',
        'measure': 'руб./м2',
        'output-data': df1.iloc[1]['output-data'],
        'arenda': df1.iloc[1]['arenda'],
        'vikup_price': df1.iloc[1]['vikup_price'],
        'arenda_price': df1.iloc[1]['arenda_price'],
        'vikup': df1.iloc[1]['vikup'],
    },
    {
        'input-data': 'Плата за технологическое присоединение к сетям',
        'measure': 'руб.',


        'output-data': df1.iloc[2]['output-data'],
        'arenda': df1.iloc[2]['arenda'],
        'vikup_price': df1.iloc[2]['vikup_price'],


    }
]
if df_help1['key'][0] == 1:
    x = data1[0]['vikup']
else:
    x = data1[0]['vikup_price']
data=[
    {

        'input-data':'Площадь на единицу производственной мощности',
        'measure': 'м2',
        '20k': df.iloc[0]['20k'],
        '40k': df.iloc[0]['40k'],
        '80k': df.iloc[0]['80k'],
        '120k': df.iloc[0]['120k'],
        '240k': df.iloc[0]['240k'],
        '360k': df.iloc[0]['360k'],

    },
    {
        'input-data':'Общая площадь',
        'measure': 'м2',
        '20k': df.iloc[1]['20k'],
        '40k': df.iloc[1]['40k'],
        '80k': df.iloc[1]['80k'],
        '120k': df.iloc[1]['120k'],
        '240k': df.iloc[1]['240k'],
        '360k': df.iloc[1]['360k'],
        'output-data': df.iloc[1]['output-data'],
        'arenda': x,   #data1[0]['vikup'] if df_help1['key'].get(0)==1 else data1[0]['vikup_price'],
        'vikup_price': df.iloc[1]['vikup_price'],
        'arenda_price': df.iloc[1]['arenda_price'],
        'vikup': df.iloc[1]['vikup'],

    },
    {
        'input-data': 'ИТОГО',


        '20k': df.iloc[2]['20k'],
        '40k': df.iloc[2]['40k'],
        '80k': df.iloc[2]['80k'],
        '120k': df.iloc[2]['120k'],
        '240k': df.iloc[2]['240k'],
        '360k': df.iloc[2]['360k'],
        'output-data': df.iloc[2]['output-data'],

        'arenda_price': df.iloc[2]['arenda_price'],
        'vikup': df.iloc[2]['vikup'],

    }
]
columnDefs=[
    {
        'headerName': 'Земля',

        'field': 'input-data',
    },
    {
        'headerName': 'Ед. Изм.', 'field': 'measure',


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
        'headerName': 'Количество, ед. изм.', 'field': 'output-data'
    },
    {
        'headerName': 'Аренда', 'field': 'arenda'
    },
    {
        'headerName': 'Выкуп, руб./м2', 'field': 'vikup_price'
    },

    {
        'headerName': 'Аренда, руб. в год', 'field': 'arenda_price'
    },
    {
        'headerName': 'Выкуп, руб.', 'field': 'vikup'
    },
]


columnDefs1=[
    { 'headerName': 'Удельная стоимость земельного участка',
      'children':[
    {
        'headerName': 'Наименование',

        'field': 'input-data',
    },
    {
        'headerName': 'Ед. Изм.', 'field': 'measure',


    },

    {
        'field': '20k', 'headerName': '',

        'width': 100,
    },
    {
        'field': '40k', 'headerName': '',

        'width': 100,
    },
    {
        'field': '80k', 'headerName': '',

        'width': 100,
    },
    {
        'field': '120k', 'headerName': '',

        'width': 100,
    },
    {
        'field': '240k', 'headerName': '',

        'width': 100,
    },
    {
        'field': '360k', 'headerName': '',

        'width': 100,
    },

    {
        'headerName': 'Средняя кадастровая стоимость ЗУ ед.изм.',
        'field': 'output-data',
        'editable':True,
    },
    {
        'headerName': 'Величина ставки от кадастровой стоимости',
        'field': 'arenda',
        'editable': True,

    },
    {
        'headerName': 'Стоимость, ед.изм.',
        'field': 'vikup_price'
    },

    {
        'headerName': 'Величина ставки от кадастровой стоимости в ОЭЗ Узловая, Тульская область',
        'field': 'arenda_price',
        'editable': True,

    },
    {
        'headerName': 'Стоимость в ОЭЗ Узловая, Тульская область, ед.изм.', 'field': 'vikup'
    },

    ]
      }
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
                'a2': str(df_help['var2'].get(df_help['key'].get(0)-1))
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
            style={"height":250},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},


            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},


        ),
        dag.AgGrid(
            style={"height":260},
            id='computed-table1',
            rowData=data1,
            columnDefs=columnDefs1,
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
    Output('computed-table', 'rowData', allow_duplicate = True),

    Input('computed-table', 'cellValueChanged'),
    State('computed-table', 'rowData'),
    State('computed-table1', 'rowData'),
    prevent_initial_call=True,

    prevent_initial_callbacks = True
)

def update_row_data_zu_1(cell_changed, data, data1):

    ind=cell_changed[0]['colId']
    data[1][ind] = data[0][ind] * dict[ind]
    data[2][ind] = data[1][ind]
    data[1]['output-data']=data[1][[key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0]]
    data[2]['output-data'] = data[1]['output-data']
    if df_help1['key'].get(0)==1:
        data[1]['arenda']=data1[0]['vikup']
    else:
        data[1]['arenda']=data1[0]['vikup_price']
    data[1]['arenda_price']=data[1]['output-data']*data[1]['arenda']
    data[2]['arenda_price']=data[1]['arenda_price']
    data[1]['vikup']=data[1]['output-data']*data[1]['vikup_price']
    data[2]['vikup']=data[1]['vikup']

    return data

@callback(

Output('computed-table', 'rowData'),
    Output('computed-table1', 'rowData'),
    Input('computed-table1', 'cellValueChanged'),
    State("computed-table", "rowData"),
    State('computed-table1', 'rowData'),

    prevent_initial_call=True,
    allow_duplicate=True,
    prevent_initial_callbacks = True
)

def update_row_data_zu_2(cell_changed, data, data1):


    ind=cell_changed[0]['rowIndex']
    data1[ind]['vikup_price'] = data1[ind]['output-data'] * data1[ind]['arenda']
    data1[ind]['vikup'] = data1[ind]['output-data'] * data1[ind]['arenda_price']
    if df_help1['key'].get(0)==1:
        data[1]['arenda']=data1[0]['vikup']
    else:
        data[1]['arenda']=data1[0]['vikup_price']
    return data, data1



@app.callback(
    Output("alerting", "is_open"),
    Output("alerting", "children"),
    Output("alerting", "color"),
    Input("save-btn", "n_clicks"),

    State("computed-table", "rowData"),
    State("computed-table1", "rowData"),
    prevent_initial_call=True,
)

def update_portfolio_stats(n, data, data1):
    dff = pd.DataFrame(data)
    dff1 = pd.DataFrame(data1)
    #df_final = pd.concat([dff, dff1])
    with pd.ExcelWriter('test.xlsx', mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
        dff.to_excel(writer, sheet_name="ЗУ_1", index=False)
        dff1.to_excel(writer, sheet_name="ЗУ_2", index=False)
    return True, "Data Saved! Well done!", "success"
'''
clientside_callback(
    """function (n) {
        if (n) {
            dash_ag_grid.getApi("computed-table").exportDataAsExcel();
        }
        return dash_clientside.no_update
    }""",
    Output("btn-excel-export", "n_clicks"),
    Input("btn-excel-export", "n_clicks"),
    prevent_initial_call=True
)'''
if __name__ == '__main__':
    app.run(debug=True)
