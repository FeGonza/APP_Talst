import dash
from dash import Dash, html, dcc , Input , Output , State, dash_table
import plotly.express as px
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta

app = Dash(__name__, use_pages=True)


# assume you have a "long-form" data frame
# see https://plotly.com/python/px-arguments/ for more options
apontamento_full = pd.read_excel(r'C:\Users\Pichau\Desktop\TALST\PROJETO_TALTS\APP_Talst\Apontamento.xlsx',sheet_name="Planilha1", index_col=None)
conferencia_full = pd.read_excel(r'C:\Users\Pichau\Desktop\TALST\PROJETO_TALTS\APP_Talst\gestta.busca.xlsx',sheet_name="Tarefas", index_col=None)

apontamento = apontamento_full.drop(['Cliente - Código', 'Usuário - Custo/Hora', 'Vencimento mês', 'Vencimento Completa'],axis = 1)
conferencia = conferencia_full.drop(['Id', 'Tipo', 'Status', 'Departamento', 'Data Legal', 'Competência', 'Feito atrasado', 'Feito com multa','Baixado'], axis = 1)

conferencia['Vencimento'] = pd.to_datetime(conferencia['Vencimento'], format='%Y-%m-%d').dt.date.astype('datetime64[ms]')
conferencia['Data de Conclusão'] = pd.to_datetime(conferencia['Data de Conclusão'], format='%Y-%m-%d').dt.date.astype('datetime64[ms]')

def transformar_horas(horas_com_ponto):
    horas=horas_com_ponto//1
    minutos = round(horas_com_ponto%1,4)*60//1
    segundos = (round(horas_com_ponto%1,4)*60%1)*60//1
    hora_transformada = str(int(horas)).zfill(2)+'h '+str(int(minutos)).zfill(2)+'min '+str(int(segundos)).zfill(2)+'seg'
    return hora_transformada
# Criando coluna de Data de Vencimento
def corrigindo_mes(mes):
    meses_do_ano=['','Janeiro','Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
    i=0
    while mes != meses_do_ano[i]:
        i+=1
    return i  
    
data_vencimento_completa=[]
for i in range(len(apontamento['Cliente - Nome'])):
    data_corrigida = date (year = apontamento['Tarefa - Data de Vencimento.Ano'][i],
                          month = corrigindo_mes(apontamento['Tarefa - Data de Vencimento.Mês'][i]), 
                          day = apontamento['Tarefa - Data de Vencimento.Dia'][i])
    data_vencimento_completa.append(data_corrigida)
                          
apontamento['Data Vencimento'] = data_vencimento_completa
apontamento['Data Vencimento'] = pd.to_datetime(apontamento['Data Vencimento'], format='%Y-%m-%d').dt.date.astype('datetime64[ms]')

data_atual = date(datetime.now().year, datetime.now().month, datetime.now().day)
start_date_default = data_atual - timedelta(30)
#Criando campo Data de Conclusão na planilha de apontamento

data_de_conclusao = []

tamanho_lista = len(apontamento['Cliente - Nome'])
for i in range(tamanho_lista):
    similaridade_tuple = np.where(conferencia['Cliente'] == apontamento['Cliente - Nome'][i])
    similaridade_array = similaridade_tuple[0]
    cont = 0
    for j in similaridade_array:
        if ((conferencia['Responsável'][j] == apontamento['Usuário - Nome'][i]) 
            and (conferencia['Nome'][j] == apontamento['Tarefa - Nome'][i]) 
            and (conferencia['Vencimento'][j] == apontamento['Data Vencimento'][i])):
                data_de_conclusao.append(conferencia['Data de Conclusão'][j])
                break
        else:
            cont+=1
    if cont == len(similaridade_array):
        data_de_conclusao.append('')
        
apontamento['Data de conclusão'] = data_de_conclusao        
apontamento['Data de conclusão'] = pd.to_datetime(apontamento['Data de conclusão'], format='%Y-%m-%d').dt.date.astype('datetime64[ms]')
apontamento = apontamento.drop(['Tarefa - Data de Vencimento.Ano', 'Tarefa - Data de Vencimento.Mês', 'Tarefa - Data de Vencimento.Dia'], axis = 1)

lista_usuarios = sorted((apontamento['Usuário - Nome'].unique()).tolist())
lista_usuarios.insert(0,'Todos')

lista_departamento = sorted((apontamento['Empresa - Departamento'].unique()).tolist())
lista_departamento.insert(0,'Todos')

image_path = 'assets/logo-talst.png'

def filtrar_departamento(df, departamento):
    if departamento == 'Todos':
        return df
    else:    
        selecionando_departamento = df.loc[(df['Empresa - Departamento']) == departamento , :]
        return selecionando_departamento    

def filtrando_usuario(df, usuario):
    if usuario == 'Todos':
        return df
    else:
        selecionando_usuario = df.loc[(df['Usuário - Nome']) == usuario , :]
        return selecionando_usuario

## Agrupando Serviços parecidos:
for tarefas in range(len(apontamento['Tarefa - Nome'])):
    if 'ADMISSÃO' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'ADMISSÃO', inplace=True)
    elif 'ALTERAÇÃO DE CARGO E SALARIO' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'ALTERAÇÃO DE CARGO E SALARIO', inplace=True)
    elif 'RESCISÃO' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'SERVIÇOS DE RESCISÃO', inplace=True)
    elif 'CONFERENCIA DE CHAMADOS' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'CONFERENCIA DE CHAMADOS', inplace=True)
    elif 'FÉRIAS CLT' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'EXECUÇÃO FÉRIAS CLT', inplace=True)
    elif 'RECALCULO DE IMPOSTO' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'RECALCULO DE IMPOSTO', inplace=True)
    elif 'FECHAMENTO FOLHA COM APONTAMENTO - DIA 5' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'FECHAMENTO FOLHA COM APONTAMENTO - 5º DIA UTIL', inplace=True)
    elif 'ALOCAÇÕES' in apontamento['Tarefa - Nome'][tarefas]:
        apontamento['Tarefa - Nome'].replace(apontamento['Tarefa - Nome'][tarefas], 'ALOCAÇÕES', inplace=True)
    else:
        sem_opcao=1 

celula_inicial = {'row': 0, 'column': 0, 'column_id': 'Tarefa - Nome', 'row_id': 0 }
celula_inicial2 = {'row': 0, 'column': 0, 'column_id': 'Tarefa - Nome', 'row_id': 1 }
# Criando HTML

app.layout = html.Div(children=[
    html.Div([
        html.Img(src=image_path, style ={'margin-top':'15px'}),
    ], style = {'display':'flex' , 
                'justify-content':'center'}
    ),

    html.Div(id='exportando', style= { 'opacity': '0'}),
    
    html.Div([
  
        html.Div([
                html.H3('Selecione o intervalo de data:'),
                dcc.DatePickerRange(
                    id='my-date-picker-range',
                    display_format='DD-MM-YYYY',
                    min_date_allowed=date(2018, 1, 1),
                    max_date_allowed=date(2024, 1, 1),
                    initial_visible_month=data_atual,
                    end_date=data_atual,
                    start_date=start_date_default,
            ),
        ], style = {'justify-content':'center' }
        ),       
    ], style= {'display':'flex' , 
                'justify-content':'space-around'}
    ),   


    dash.page_container,
    

], style = {'background-image':'linear-gradient(89deg, #05ff8b3d, #005ce53d)', 
            'min-height':'100vh'}
)




# PAGINA home (TABELA DE TAREFAS): 

## CRIANDO PLANILHA DE EXIBIÇAO 

@app.callback(
    Output(component_id='tabela_tarefas' , component_property= 'data'),
    [Input(component_id='lista de departamento' , component_property='value'),
     Input(component_id='lista de usuarios' , component_property='value'),
     Input(component_id='my-date-picker-range' , component_property='start_date'),
     Input(component_id='my-date-picker-range' , component_property='end_date')]  
)
def update_tabela_tarefas(value_departamento, value_usuario, start_date, end_date ):
    apontamento_copy = apontamento.copy()
    tabela_filtrada_departamento = filtrar_departamento(apontamento_copy, value_departamento)
    tabela_filtrada_usuario = filtrando_usuario(tabela_filtrada_departamento, value_usuario)
    data_selecao = tabela_filtrada_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
    tabela_filtrada = data_selecao.groupby('Tarefa - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
    if len(tabela_filtrada) == 0:
        data_default = [['NAO POSSUE TAREFAS NESSE PERIODO', 0.0]]
        tabela_filtrada = pd.DataFrame(data_default, columns=['Tarefa - Nome', 'Apontamento - Tempo (em horas)'])
    else:        
        tabela_filtrada.dropna(subset=['Apontamento - Tempo (em horas)'])
        tabela_filtrada['Apontamento - Tempo (em horas)'] = list(map(transformar_horas, tabela_filtrada['Apontamento - Tempo (em horas)']))
    return (tabela_filtrada.to_dict('records'))

@app.callback(
    Output(component_id='lista de departamento' , component_property= 'options'),
    Input(component_id='exportando' , component_property='value')
)
def update_dropdown_lista_de_departamento(value):
    lista_departamento = sorted((apontamento['Empresa - Departamento'].unique()).tolist())
    lista_departamento.insert(0,'Todos')
    return lista_departamento

@app.callback(
    Output(component_id='lista de usuarios' , component_property= 'options'),
    Input(component_id='lista de departamento' , component_property='value')
)
def update_dropdown_lista_de_usuario(value):
    apontamento_copy = apontamento.copy()
    separando_por_departamento = filtrar_departamento(apontamento_copy, value)
    lista_usuarios = sorted((separando_por_departamento['Usuário - Nome'].unique()).tolist())
    lista_usuarios.insert(0,'Todos')
    return (lista_usuarios)


## CRINADO BOTAO PARA EXPORTAR PARA EXCEL

@app.callback(
    Output(component_id='download-dataframe-xlsx', component_property='data'),
    Input(component_id='btn_xlsx', component_property='n_clicks'),
    [State(component_id='lista de departamento' , component_property='value'),
     State(component_id='lista de usuarios' , component_property='value'),
     State(component_id='my-date-picker-range' , component_property='start_date'),
     State(component_id='my-date-picker-range' , component_property='end_date')],
    prevent_initial_call=True,
)
def funcao_botao_export_excel(n_clicks, value_departamento,  value_usuario, start_date, end_date):
    if n_clicks > 0:
        if value_usuario == 'Todos':
            apontamento_copy = apontamento.copy()
            tabela_filtrada_departamento = filtrar_departamento(apontamento_copy, value_departamento)
            tabela_filtrada_usuario = filtrando_usuario(tabela_filtrada_departamento, value_usuario)
            data_selecao = tabela_filtrada_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
            tabela_filtrada = data_selecao.groupby('Tarefa - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
            tabela_filtrada.rename(columns={'Tarefa - Nome':'Tarefa - Nome (TODOS)'})
            return dcc.send_data_frame(tabela_filtrada.to_excel, "TalstResumo.xlsx", sheet_name="planilha1")
        else:    
            apontamento_copy = apontamento.copy()
            tabela_filtrada_departamento = filtrar_departamento(apontamento_copy, value_departamento)
            tabela_filtrada_usuario = filtrando_usuario(tabela_filtrada_departamento, value_usuario)
            data_selecao = tabela_filtrada_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
            tabela_filtrada = data_selecao.groupby('Tarefa - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
            tabela_filtrada.rename(columns={'Tarefa - Nome' : f'Tarefa - Nome ({value_usuario})'})
            return dcc.send_data_frame(tabela_filtrada.to_excel, "TalstResumo.xlsx", sheet_name="planilha1")

## BOTAO PARA ACESSAR OUTRA PAGINA

@app.callback(
    Output(component_id='exportando', component_property='children'),
    Input(component_id='tabela_tarefas' , component_property='selected_rows'),
    [State(component_id='lista de departamento' , component_property='value'),
     State(component_id='lista de usuarios' , component_property='value'),
     State(component_id='my-date-picker-range' , component_property='start_date'),
     State(component_id='my-date-picker-range' , component_property='end_date')],
    prevent_initial_call=True,
)
def funcao_exportando_info(selected_rows, value_departamento, value_usuario, start_date, end_date):
    if len(selected_rows) == 0:
        return 'Selecione uma Tarefa'
    else:
        if value_usuario == 'Todos':
            apontamento_copy = apontamento.copy()
            tabela_filtrada_departamento = filtrar_departamento(apontamento_copy, value_departamento)
            tabela_filtrada_usuario = filtrando_usuario(tabela_filtrada_departamento, value_usuario)
            data_selecao = tabela_filtrada_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
            tabela_filtrada = data_selecao.groupby('Tarefa - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
            selecao = selected_rows[0]
            return tabela_filtrada['Tarefa - Nome'][selecao]
        else:    
            apontamento_copy = apontamento.copy()
            tabela_filtrada_departamento = filtrar_departamento(apontamento_copy, value_departamento)
            tabela_filtrada_usuario = filtrando_usuario(tabela_filtrada_departamento, value_usuario)
            data_selecao = tabela_filtrada_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
            tabela_filtrada = data_selecao.groupby('Tarefa - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
            if len(tabela_filtrada) == 0:
                return ' '
            else:    
                selecao = selected_rows[0]
                return tabela_filtrada['Tarefa - Nome'][selecao]

# PAGINA resumotarefas (TABELA DE COLABORADORES): 

## CRIANDO TABELA 

@app.callback(
    Output(component_id='tabela_colaboradores' , component_property= 'data'),
    [Input(component_id='exportando', component_property='children'),
     Input(component_id='my-date-picker-range' , component_property='start_date'),
     Input(component_id='my-date-picker-range' , component_property='end_date')]  
)
def update_tabela_colaboradores(tarefa, start_date, end_date ): 
    apontamento_copy = apontamento.copy()
    selecionando_usuario = apontamento_copy.loc[(apontamento_copy['Tarefa - Nome']) == tarefa , :]
    data_selecao = selecionando_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
    tabela_filtrada = data_selecao.groupby('Usuário - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
    if len(tabela_filtrada) == 0:
        data_default = [['NAO POSSUE TAREFAS NESSE PERIODO', 0.0]]
        tabela_filtrada = pd.DataFrame(data_default, columns=['Tarefa - Nome', 'Apontamento - Tempo (em horas)'])
    else:        
        tabela_filtrada.dropna(subset=['Apontamento - Tempo (em horas)'])
        tabela_filtrada['Apontamento - Tempo (em horas)'] = list(map(transformar_horas, tabela_filtrada['Apontamento - Tempo (em horas)']))
    return (tabela_filtrada.to_dict('records'))

@app.callback(
    Output(component_id='grafico_colaboradores' , component_property= 'figure'),
    [Input(component_id='exportando', component_property='children'),
     Input(component_id='my-date-picker-range' , component_property='start_date'),
     Input(component_id='my-date-picker-range' , component_property='end_date')]  
)
def update_grafico_colaboradores(tarefa, start_date, end_date ): 
    apontamento_copy = apontamento.copy()
    selecionando_usuario = apontamento_copy.loc[(apontamento_copy['Tarefa - Nome']) == tarefa , :]
    data_selecao = selecionando_usuario.loc[(apontamento['Data de conclusão'] >= np.datetime64(start_date) ) & (apontamento['Data de conclusão'] <= np.datetime64(end_date) )]
    tabela_filtrada = data_selecao.groupby('Usuário - Nome').sum().sort_values(by=['Apontamento - Tempo (em horas)'], ascending=False).reset_index()
    
    if len(tabela_filtrada) == 0:
        data_default = [['NAO POSSUE TAREFAS NESSE PERIODO', 0.0]]
        piechart=px.bar(
            data_frame=data_default, 
            x = 'Usuário - Nome',
            y = 'Apontamento - Tempo (em horas)'
        )
        return (piechart)
    else:
        tabela_filtrada.dropna(subset=['Apontamento - Tempo (em horas)'])
        piechart=px.bar(
            data_frame=data_default, 
            x = 'Usuário - Nome',
            y = 'Apontamento - Tempo (em horas)'
        )
        return (piechart)


@app.callback(
    Output(component_id='tipo_tarefa', component_property='children'),
    Input(component_id='exportando', component_property='children')
)
def exibir_tarefa(children):
    return f'TAREFA: {children}'

if __name__ == '__main__':
    app.run_server(debug=True)

