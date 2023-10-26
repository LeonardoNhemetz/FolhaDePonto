import gspread
import pandas as pd
import secret
# Autenticação e obtenção de dados do Google Planilhas

gc = gspread.service_account(filename='key.json')
sh = gc.open_by_key(secret.planCode)
ws = sh.worksheet('Outubro')

# Obtendo dados do Google Planilhas e criando DataFrame
dados = pd.DataFrame(ws.get_all_records())

# Removendo linhas que contêm valores "NaN" nas colunas necessárias
dados = dados.dropna(subset=['Entrada', 'Saida', 'Entrada.AL', 'Saida.AL'])

# Convertendo as colunas 'Entrada', 'Saida', 'Entrada.AL' e 'Saida.AL' para o formato datetime
dados['Entrada'] = pd.to_datetime(
    dados['Entrada'], format='%H:%M', errors='coerce')
dados['Saida'] = pd.to_datetime(
    dados['Saida'], format='%H:%M', errors='coerce')
dados['Entrada.AL'] = pd.to_datetime(
    dados['Entrada.AL'], format='%H:%M', errors='coerce')
dados['Saida.AL'] = pd.to_datetime(
    dados['Saida.AL'], format='%H:%M', errors='coerce')

# Calculando as horas trabalhadas e o horário de almoço
dados['HorasTrabalhadas'] = (
    (dados['Saida'] - dados['Entrada']).dt.total_seconds() / 3600).apply(lambda x: round(x, 2))
dados['HorarioAlmoco'] = (
    (dados['Entrada.AL'] - dados['Saida.AL']).dt.total_seconds() / 3600).apply(lambda x: round(x, 2))

# Função para formatar horas em H:M


def formatar_horas(horas):
    if pd.notnull(horas):
        horas_inteiras = int(horas // 1)  # Obtendo a parte inteira das horas
        # Convertendo a parte decimal em minutos
        minutos = int((horas % 1) * 60)
        return f'{horas_inteiras:02d}:{minutos:02d}'
    else:
        return "-"


# Aplicando a função de formatação às colunas 'HorasTrabalhadas' e 'HorarioAlmoco'
dados['Horas_Totais'] = dados['HorasTrabalhadas'].apply(formatar_horas)
dados['Horario_Almoco_Total'] = dados['HorarioAlmoco'].apply(formatar_horas)

# Calculando o total de horas trabalhadas subtraindo o horário do almoço
dados['HorasTrabalhadasSemAlmoco'] = dados['HorasTrabalhadas'] - \
    dados['HorarioAlmoco']

# Formatando 'HorasTrabalhadasSemAlmoco' para H:M
dados['Horas_Trabalhadas'] = dados['HorasTrabalhadasSemAlmoco'].apply(
    formatar_horas)

# Selecionando todas as colunas necessárias
resultado = dados[['Data', 'Dia Semana', 'Horas_Totais',
                   'Horario_Almoco_Total', 'Horas_Trabalhadas']]

# Mostrando o resultado final
print(resultado)

# Somando todas as horas trabalhadas
total_horas_trabalhadas = dados['HorasTrabalhadasSemAlmoco'].sum()
print(f"Total de horas trabalhadas: {total_horas_trabalhadas:.2f} horas")
