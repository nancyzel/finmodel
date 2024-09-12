import dash
from dash import Dash, dash_table, html, Input, Output, State, callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc


import pandas as pd



df=pd.read_excel('test.xlsx', sheet_name='ФЗП')
df_help=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df_help1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_энергцентр')
df_help2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторферм')
df_help3=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторсеп')
df_help4=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторсуш')


app = Dash(__name__)

data=[{
    'input-data':row.iloc[0],
    'category': row.iloc[1],
    'sep': row.iloc[2],
    'expend': row.iloc[3],
    'sh1_1': row.iloc[4],
    'sh2_1': row.iloc[5],
    'all1': row.iloc[6],
    'sh1_2': row.iloc[7],
    'sh2_2': row.iloc[8],
    'all2': row.iloc[9],
    'sh1_3': row.iloc[10],
    'sh2_3': row.iloc[11],
    'all3': row.iloc[12],
    'sh1_4': row.iloc[13],
    'sh2_4': row.iloc[14],
    'all4': row.iloc[15],
    'sh1_5': row.iloc[16],
    'sh2_5': row.iloc[17],
    'all5': row.iloc[18],
    'sh1_6': row.iloc[19],
    'sh2_6': row.iloc[20],
    'all6': row.iloc[21],
    'many': row.iloc[22],
    'tgp': row.iloc[23],
    'month': row.iloc[24],
    'summam': row.iloc[25],
    'summay': row.iloc[26],
} for ind, row in df.iterrows()]


columnDefs=[
    {
        'headerName': 'Наименование должности',

        'field': 'input-data',
        'editable': True,
    },
    {
        'headerName': 'Категория персонала',

        'field': 'category',
        'editable': True,

    },
    {
        'headerName': 'Подразделение',

        'field': 'sep',
        'editable': True,

    },
    {
        'headerName': 'Вид затрат',
        'field': 'expend',
        'editable': True,

    },
    {
        'headerName': 'Производственная мощность в год, тонн',
        'children':[
            {
                'headerName': '20000',
                'editable':True,
                'children': [
                    {
                        'field':'sh1_1',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_1',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all1',
                        'headerName': 'Всего'
                    },
                ]
            },
            {
                'headerName': '40000',
                'editable':True,
                'children': [
                    {
                        'field':'sh1_2',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_2',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all2',
                        'headerName': 'Всего'
                    },
                ]
            },
            {
                'headerName': '80000',
                'editable': True,
                'children': [
                    {
                        'field': 'sh1_3',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_3',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all3',
                        'headerName': 'Всего'
                    },
                ]
            },
            {
                'headerName': '120000',
                'editable': True,
                'children': [
                    {
                        'field': 'sh1_4',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_4',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all4',
                        'headerName': 'Всего'
                    },
                ]
            },            {
                'headerName': '240000',
                'editable':True,
                'children': [
                    {
                        'field':'sh1_5',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_5',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all5',
                        'headerName': 'Всего'
                    },
                ]
            },
            {
                'headerName': '480000',
                'editable':True,
                'children': [
                    {
                        'field':'sh1_6',
                        'headerName': 'смена 1 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'sh2_6',
                        'headerName': 'смена 2 (12 ч)',
                        'editable': True
                    },
                    {
                        'field': 'all6',
                        'headerName': 'Всего'
                    },
                ]
            },
        ]

    },
    {
        'headerName': 'Численность',

        'field': 'many',

    },

    {
        'headerName': 'Ставка, руб./т.г.п.',
        'field': 'tgp',
    },
    {
        'headerName': 'Ставка, руб./мес.',
        'field': 'month',

    },
    {
        'headerName': 'Сумма, руб./мес.', 'field': 'summam'
    },
    {
        'headerName': 'Сумма, руб./год', 'field': 'summay'
    },
]



app.layout = html.Div(
    [
        dag.AgGrid(
            style={"height":200, "width":"100%"},
            id='small-table',
            dashGridOptions = {'suppressNoRowsOverlay':True},
            rowData=[
                {
                'a1': 'Производственная мощность в год, тонн',
                'a2': str(df_help['var2'].get(df_help['key'].get(0)-1)),
                },
                {
                    'a1': 'Строительство собственного энергоцентра',
                    'a2': df_help1['var'].get(df_help1['key'].get(0) - 1),
                },

            ],
            columnDefs=[
                {
                    'headerName': 'Ставки оплаты труда',
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
            style={"height":2800},
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
sp_sovm=['all1','all2','all3','all4','all5','all6']
alls=['sh1_1', 'sh2_1',
      'sh1_2', 'sh2_2',
      'sh1_3', 'sh2_3',
      'sh1_4', 'sh2_4',
      'sh1_5', 'sh2_5',
      'sh1_6', 'sh2_6']
alls_dict={
    'sh1_1': 'all1', 'sh2_1':'all1',
    'sh1_2':'all2', 'sh2_2':'all2',
    'sh1_3': 'all3', 'sh2_3': 'all3',
    'sh1_4': 'all4', 'sh2_4': 'all4',
    'sh1_5': 'all5', 'sh2_5': 'all5',
    'sh1_6': 'all6', 'sh2_6': 'all6'
}
b2=df_help['var2'].get(df_help['key'].get(0)-1)

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
    if ind in alls:
        data[indr][alls_dict[ind]]=data[indr][ind]+data[indr][alls[alls.index(ind)+(-1)**alls.index(ind)%2!=0]]
        for i in range(6):
            data[58][alls[2*i]]=1 if df_help1['var'].get(df_help1['key'].get(0) - 1)=='Да' else 0
            data[60][alls[2*i]]=1 if df_help1['var'].get(df_help1['key'].get(0) - 1)=='Да' else 0
            data[60][alls[2*i+1]]=1 if df_help1['var'].get(df_help1['key'].get(0) - 1)=='Да' else 0

        data[indr]['many']=data[indr][sp_sovm[sp.index([key for key, value in dict.items() if value==df_help['var2'].get(df_help['key'].get(0)-1)][0])]]
        data[29]['tgp']=df_help2['var'].get(df_help1['key'].get(0) - 1)
        data[30]['tgp']=df_help2['var'].get(df_help1['key'].get(0) - 1)
        data[31]['tgp']=df_help3['var'].get(df_help1['key'].get(0) - 1)
        data[32]['tgp']=df_help4['var'].get(df_help1['key'].get(0) - 1)
        if data[indr]['tgp']==0:
            data[indr]['summam']=data[indr]['many']*data[indr]['month']
        else:
            data[indr]['summam']=data[indr]['many']*data[indr]['tgp']*b2/12
        data[indr]['summay']=data[indr]['summam']*12
        for i in range(29,33):
            data[i]['month']=data[i]['summam']/data[i]['many']
    elif ind in ['tgp', 'month']:
        data[29]['tgp']=df_help2['var'].get(df_help1['key'].get(0) - 1)
        data[30]['tgp']=df_help2['var'].get(df_help1['key'].get(0) - 1)
        data[31]['tgp']=df_help3['var'].get(df_help1['key'].get(0) - 1)
        data[32]['tgp']=df_help4['var'].get(df_help1['key'].get(0) - 1)
        if data[indr]['tgp']==0:
            data[indr]['summam']=data[indr]['many']*data[indr]['month']
        else:
            data[indr]['summam']=data[indr]['many']*data[indr]['tgp']*b2/12
        data[indr]['summay']=data[indr]['summam']*12
        for i in range(29,33):
            data[i]['month']=data[i]['summam']/data[i]['many']
    data[64]['summam']=sum([data[j]['summam'] for j in range(64)])
    data[64]['summay']=sum([data[j]['summay'] for j in range(64)])

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
        dff.to_excel(writer, sheet_name="ФЗП", index=False)
    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)