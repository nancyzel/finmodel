import dash
from dash import Dash, dash_table, html, Input, Output, State, callback, clientside_callback
import dash_ag_grid as dag
import dash_bootstrap_components as dbc

#import plotly.express as px
import pandas as pd
#import js2py

#js_add='''function isCellEditable('''
df=pd.read_excel('test.xlsx', sheet_name='СлужСпр_производмощ')
df1=pd.read_excel('test.xlsx', sheet_name='СлужСпр_земуч')
df2=pd.read_excel('test.xlsx', sheet_name='СлужСпр_энергцентр')
df3=pd.read_excel('test.xlsx', sheet_name='СлужСпр_обеспечкислород')
df4=pd.read_excel('test.xlsx', sheet_name='СлужСпр_лизинг')
df5=pd.read_excel('test.xlsx', sheet_name='СлужСпр_электрмощ')
df6=pd.read_excel('test.xlsx', sheet_name='СлужСпр_странаизг')
df7=pd.read_excel('test.xlsx', sheet_name='СлужСпр_доппрод')
df8=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторферм')
df9=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторсеп')
df10=pd.read_excel('test.xlsx', sheet_name='СлужСпр_операторсуш')
df11=pd.read_excel('test.xlsx', sheet_name='СлужСпр_рабочдавл')
df12=pd.read_excel('test.xlsx', sheet_name='СлужСпр_запасмтр')
df13=pd.read_excel('test.xlsx', sheet_name='СлужСпр_льгсубс')
df14=pd.read_excel('test.xlsx', sheet_name='СлужСпр_парампроизв')
df15=pd.read_excel('test.xlsx', sheet_name='СлужСпр_колвоферм')
df16=pd.read_excel('test.xlsx', sheet_name='СлужСпр_схемафинанс')

#print(df.to_string())

app = Dash(__name__)

data=[{
    'ind':row.iloc[0],
    'var1': row.iloc[1],
    'var2': row.iloc[2],
    'key': row.iloc[3]
} for a, row in df.iterrows()]
columnDefs=[
    {
        'headerName': 'Производственная мощность в год',
        'field': 'ind',
    },
    {
        'headerName': '', 'field': 'var1',
    },
    {
        'headerName': '', 'field': 'var2',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data1=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df1.iterrows()]
columnDefs1=[
    {
        'headerName': 'Земельный участок',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data2=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df2.iterrows()]
columnDefs2=[
    {
        'headerName': 'Энергетический центр',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data3=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df3.iterrows()]
columnDefs3=[
    {
        'headerName': 'Вариант обеспечения кислородом',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data4=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df4.iterrows()]
columnDefs4=[
    {
        'headerName': 'Лизинг оборудования',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data5=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df5.iterrows()]
columnDefs5=[
    {
        'headerName': 'Электрическая мощность',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data6=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df6.iterrows()]
columnDefs6=[
    {
        'headerName': 'Страна изготовления оборудования',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data7=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df7.iterrows()]
columnDefs7=[
    {
        'headerName': 'Дополнительная продукция',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data8=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df8.iterrows()]
columnDefs8=[
    {
        'headerName': 'Сдельная ставка Оператор ферментера',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data9=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df9.iterrows()]
columnDefs9=[
    {
        'headerName': 'Сдельная ставка Оператор сепаратора и ВВУ',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data10=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df10.iterrows()]
columnDefs10=[
    {
        'headerName': 'Сдельная ставка Оператор сушильной установки',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data11=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df11.iterrows()]
columnDefs11=[
    {
        'headerName': 'Рабочее давление в ферментере',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data12=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key1': row.iloc[2],
    'key2': row.iloc[3],
} for a, row in df12.iterrows()]
columnDefs12=[
    {
        'headerName': 'Формирование запасов МТР',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key1', 'editable': True,
    },
    {
        'headerName': '', 'field': 'key2', 'editable': True,
    },
]

data13=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df13.iterrows()]
columnDefs13=[
    {
        'headerName': 'Льготы и субсидии',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data14=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df14.iterrows()]
columnDefs14=[
    {
        'headerName': 'Параметры производства',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data15=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df15.iterrows()]
columnDefs15=[
    {
        'headerName': 'Количество ферментеров',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

data16=[{
    'ind':row.iloc[0],
    'var': row.iloc[1],
    'key': row.iloc[2]
} for a, row in df16.iterrows()]
columnDefs16=[
    {
        'headerName': 'Схема финансирования №1',
        'field': 'ind',
    },

    {
        'headerName': '', 'field': 'var',
        'editable': True,
    },
    {
        'headerName': '', 'field': 'key', 'editable': True,
    },
]

app.layout = html.Div(
    [

        dag.AgGrid(
            style={"height":350, "width": '60%'},
            id='computed-table',
            rowData=data,
            columnDefs=columnDefs,
            defaultColDef={"sortable":False},


            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},
        ),
        dag.AgGrid(
            style={"height":150, "width": '60%'},
            id='computed-table1',
            rowData=data1,
            columnDefs=columnDefs1,
            defaultColDef={"sortable":False},
            dashGridOptions={
                "suppressRowTransform":True,
                "defaultExcelExportParams": {"headerRowHeight": 30},},


        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table2',
            rowData=data2,
            columnDefs=columnDefs2,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table3',
            rowData=data3,
            columnDefs=columnDefs3,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table4',
            rowData=data4,
            columnDefs=columnDefs4,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 350, "width": '60%'},
            id='computed-table5',
            rowData=data5,
            columnDefs=columnDefs5,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 210, "width": '60%'},
            id='computed-table6',
            rowData=data6,
            columnDefs=columnDefs6,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table7',
            rowData=data7,
            columnDefs=columnDefs7,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 370, "width": '60%'},
            id='computed-table8',
            rowData=data8,
            columnDefs=columnDefs8,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 370, "width": '60%'},
            id='computed-table9',
            rowData=data9,
            columnDefs=columnDefs9,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 370, "width": '60%'},
            id='computed-table10',
            rowData=data10,
            columnDefs=columnDefs10,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 550, "width": '60%'},
            id='computed-table11',
            rowData=data11,
            columnDefs=columnDefs11,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table12',
            rowData=data12,
            columnDefs=columnDefs12,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table13',
            rowData=data13,
            columnDefs=columnDefs13,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table14',
            rowData=data14,
            columnDefs=columnDefs14,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 600, "width": '60%'},
            id='computed-table15',
            rowData=data15,
            columnDefs=columnDefs15,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),
        dag.AgGrid(
            style={"height": 150, "width": '60%'},
            id='computed-table16',
            rowData=data16,
            columnDefs=columnDefs16,
            defaultColDef={"sortable": False},
            dashGridOptions={
                "suppressRowTransform": True,
                "defaultExcelExportParams": {"headerRowHeight": 30}, },

        ),


        dbc.Col
        (
            [
                dbc.Button(

                    id="save-btn",
                    children="Save Tables",
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




@app.callback(
    Output("alerting", "is_open"),
    Output("alerting", "children"),
    Output("alerting", "color"),
    Input("save-btn", "n_clicks"),

    State("computed-table", "rowData"),
    State("computed-table1", "rowData"),
    State("computed-table2", "rowData"),
    State("computed-table3", "rowData"),
    State("computed-table4", "rowData"),
    State("computed-table5", "rowData"),
    State("computed-table6", "rowData"),
    State("computed-table7", "rowData"),
    State("computed-table8", "rowData"),
    State("computed-table9", "rowData"),
    State("computed-table10", "rowData"),
    State("computed-table11", "rowData"),
    State("computed-table12", "rowData"),
    State("computed-table13", "rowData"),
    State("computed-table14", "rowData"),
    State("computed-table15", "rowData"),
    State("computed-table16", "rowData"),

    prevent_initial_call=True,
)

def update_portfolio_stats(n, data, data1, data2, data3,data4,data5,data6,data7,data8,data9,data10,data11,data12,data13,data14,data15,data16):
    dff = pd.DataFrame(data)
    dff1 = pd.DataFrame(data1)
    dff2 = pd.DataFrame(data2)
    dff3 = pd.DataFrame(data3)
    dff4 = pd.DataFrame(data4)
    dff5 = pd.DataFrame(data5)
    dff6 = pd.DataFrame(data6)
    dff7 = pd.DataFrame(data7)
    dff8 = pd.DataFrame(data8)
    dff9 = pd.DataFrame(data9)
    dff10 = pd.DataFrame(data10)
    dff11 = pd.DataFrame(data11)
    dff12 = pd.DataFrame(data12)
    dff13 = pd.DataFrame(data13)
    dff14 = pd.DataFrame(data14)
    dff15 = pd.DataFrame(data15)
    dff16 = pd.DataFrame(data16)

    with pd.ExcelWriter('test.xlsx', mode="a", engine="openpyxl", if_sheet_exists='replace') as writer:
        dff.to_excel(writer, sheet_name="СлужСпр_производмощ", index=False)
        dff1.to_excel(writer, sheet_name="СлужСпр_земуч", index=False)
        dff2.to_excel(writer, sheet_name="СлужСпр_энергцентр", index=False)
        dff3.to_excel(writer, sheet_name="СлужСпр_обеспечкислород", index=False)
        dff4.to_excel(writer, sheet_name="СлужСпр_лизинг", index=False)
        dff5.to_excel(writer, sheet_name="СлужСпр_электрмощ", index=False)
        dff6.to_excel(writer, sheet_name="СлужСпр_странаизг", index=False)
        dff7.to_excel(writer, sheet_name="СлужСпр_доппрод", index=False)
        dff8.to_excel(writer, sheet_name="СлужСпр_операторферм", index=False)
        dff9.to_excel(writer, sheet_name="СлужСпр_операторсеп", index=False)
        dff10.to_excel(writer, sheet_name="СлужСпр_операторсуш", index=False)
        dff11.to_excel(writer, sheet_name="СлужСпр_рабочдавл", index=False)
        dff12.to_excel(writer, sheet_name="СлужСпр_запасмтр", index=False)
        dff13.to_excel(writer, sheet_name="СлужСпр_льгсубс", index=False)
        dff14.to_excel(writer, sheet_name="СлужСпр_парампроизв", index=False)
        dff15.to_excel(writer, sheet_name="СлужСпр_колвоферм", index=False)
        dff16.to_excel(writer, sheet_name="СлужСпр_схемафинанс", index=False)

    return True, "Data Saved! Well done!", "success"

if __name__ == '__main__':
    app.run(debug=True)