import dash
from dash import dcc, html, dash_table 
import plotly.express as px

dash.register_page(__name__, '/resumotarefa')

layout = html.Div(children=[

        html.Div([
            html.Div([
                dcc.Link('CLIQUE AQUI PARA "HORAS UTILIZADAS x TAREFA"', 
                        href='/', 
                        style={'color': 'white', 
                                'text-decoration': 'unset'}
                )
            ], style= { 'border': '2px solid #938585',
                            'background': '#3282bca3',
                            'height': '50px',
                            'align-items': 'center',
                            'justify-content': 'center',
                            'display': 'flex',
                            'width': '450px',
                            'border-radius':'18px'}
            )
        ], style = {'display':'flex',
            'flex-direction': 'column',
            'align-items':'center',
            'margin-top':'20px' }
        ),

        
        html.H2('RESUMO DE HORAS UTILIZADAS x COLABORADOR', style={'display':'flex', 
                                                                'justify-content':'center'}),
        html.Div(id='tipo_tarefa', style={'display':'flex', 
                                           'justify-content':'center',
                                           'margin-bottom':'20px'}),
        html.Div(
            dcc.Graph(id='grafico_colaboradores',
                    figure = {'layout': {
                                        'plot_bgcolor':'rgb(0 0 0 / 0%)',
                                        'paper_bgcolor':'rgb(0 0 0 / 0%)'}
                    },
                    style = {'background-color':'rgb(0 0 0 / 0%)'}
        ), style = {'width': '700px',
                    'margin-left': 'auto',
                    'margin-right': 'auto',
                    'margin-bottom':'20px'}
        ),

        html.Div(
                dash_table.DataTable(id='tabela_colaboradores',
                                    style_cell = {'textAlign': 'left'}, 
                                    style_data = {'height': 'auto', 
                                                    'Width': '50%'},
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
    ]
)

