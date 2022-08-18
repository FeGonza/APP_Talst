import dash
from dash import dcc, html, dash_table

dash.register_page(__name__, '/')

celula_inicial = {'row': 0, 'column': 0, 'column_id': 'Tarefa - Nome', 'row_id': 0 }



layout = html.Div([   
    html.Div([
        html.Div([
            html.Div([
                html.H3('Escolha o departamento:'),
                dcc.Dropdown(
                            value = 'Todos',
                            id='lista de departamento',
                            style = {'height':'46px', 'font-size':'19px'}),
                ], style = {'width': '300px','align-itens':'center'}
            ),   
            html.Div([
                html.H3('Escolha o colaborador:'),
                dcc.Dropdown(value = 'Todos',
                            id='lista de usuarios',
                            style = {'height':'46px', 'font-size':'19px'}),
                ], style = {'width': '300px','align-itens':'center'}
            ),
        ], style = {'display':'flex',
                    'margin-bottom':'15px',
                    'width': '60%',
                    'justify-content':'space-between'}),
                  

        html.Div(
                [dcc.Link('CLIQUE AQUI PARA RESUMO TAREFA SELECIONADA', 
                        href='/resumotarefa', style={'color': 'white', 
                                                    'text-decoration': 'unset'}
                )
                ], style= { 'border': '2px solid #938585',
                            'background': 'rgb(122 183 164)',
                            'height': '50px',
                            'align-items': 'center',
                            'justify-content': 'center',
                            'display': 'flex',
                            'width': '600px',
                            'border-radius':'18px'}
        )
    ], style = {'display':'flex',
                'flex-direction': 'column',
                'align-items':'center'}
    ),


        html.H2('RESUMO DE HORAS UTILIZADAS x TAREFA', style={'display':'flex', 
                                            'justify-content':'center'}),


        html.Div(
                dash_table.DataTable(id='tabela_tarefas',
                                    row_selectable='single',
                                    selected_rows=[],
                                    active_cell = celula_inicial,
                                    style_cell = {'textAlign': 'left'}, 
                                    style_data = {'height': 'auto', 
                                                    'minWidth': '50%'},
                                    style_header = {'font-weight': 'bold',
                                                    'textAlign': 'center'},
                                    style_table = {'width': '60%',
                                                    'display':'block',
                                                    'margin-left':'auto',
                                                    'margin-right':'auto'},
                                    style_cell_conditional=[{
                                                                'if': {'column_id': 'Apontamento - Tempo (em horas)'},
                                                                'textAlign': 'center'
                                                            }]
                                    )
        ),  
        
        html.Div([
            html.Button('Excel', id ='btn_xlsx', style = {'color': '#42bb24',
                                                'box-shadow': '#42bb24' ,
                                                'border': 'solid 1px #42bb24',
                                                'border-radius': '36px',
                                                'width': '65px',
                                                'height': '65px',
                                                'font-size': '20px',
                                                'cursor':'pointer'},
            ),
            dcc.Download(id="download-dataframe-xlsx"),
        ], style = {'position':'fixed', 'right':'35px' , 
                    'bottom':'100px', 'cursor': 'pointer', 
                    'height':'70px', 'width': '65px'}
        ),
    ]
)