import pandas as pd

#Usando outra função porque o caralho do pandas não está lendo o nome das abas
nomeAbas = pd.ExcelFile('Controle de Horario.xlsx')
nomeAbas = nomeAbas.sheet_names
print(f'Pessoas identificadas: {nomeAbas}')

for aba in nomeAbas:
    try:
        #Lendo a aba respectiva e criando as variaveis semana e dia
        df = pd.read_excel('Controle de Horario.xlsx', sheet_name = aba)        
        df['Semana'] = df['Data'].dt.isocalendar().week
        df['diaSemana'] = df['Data'].dt.isocalendar().day

        #Lidando com folgas
        colunas = ['Entrada', 'Almoço', 'Almoço.1', 'Saida', 'Total']
        df[colunas] = df[colunas].map(lambda x: '00:00:00' if isinstance(x, str) and x.lower() == 'folga' else x)

        #Filtrando as colunas e retirando linhas string
        horas = df[['Data','diaSemana','Semana','Total']].copy()
        horas["Total"] = horas["Total"].astype(str)
        mascara = horas["Total"].str.contains("[a-zA-Z]")
        horas = horas.loc[~mascara, :]
        semanas = horas['Semana'].unique()

        #Convertendo tudo em minutos 
        horasPicado = pd.DataFrame()
        horasPicado[['Horas', 'Minutos', '_']] = horas['Total'].str.split(":", expand=True)
        horas = pd.merge(horas, horasPicado, left_index=True, right_index=True)
        horas['Horas'] = pd.to_numeric(horas['Horas'], errors='coerce').fillna(0)
        horas['Minutos'] = pd.to_numeric(horas['Minutos'], errors='coerce').fillna(0)
        horas['MinutosTotais'] = (horas['Horas'].astype(float)*60) + (horas['Minutos'].astype(float))
        horas = horas.drop(axis=1, columns=['Horas', 'Minutos', '_'])

        #Descontando as horas que deveriam ser trabalhadas
        deusPermite = ['1', '2', '3', '4', '5', '6']
        deusNaoPermite = ['7']
        horas['diaSemana'] = horas['diaSemana'].astype(str)
        diasTrabalhados = horas.groupby('Semana')['diaSemana'].nunique().reset_index()

        for index, row in diasTrabalhados.iterrows():
            semana = row['Semana']
            diaTrabalho = row['diaSemana']
            totalMinutosSemana = horas.loc[horas['Semana'] == semana, 'MinutosTotais'].sum()

            #Se a pessoa trabalhou mais de 5 dias na semana, no sexto dia não subtrai 528
            if diaTrabalho > 5 and totalMinutosSemana > (5 * 44 * 60):
                horas.loc[(horas['Semana'] == semana) & (horas['diaSemana'].isin(deusPermite)), 'MinutosTotais'] -= 2640 / diaTrabalho

            #Se a pessoa trabalhou 5 dias ou menos na semana, subtrai 528 de cada dia
            else:
                horas.loc[(horas['Semana'] == semana) & (horas['diaSemana'].isin(deusPermite)), 'MinutosTotais'] -= 528
                horas.loc[(horas['Semana'] == semana) & (horas['diaSemana'].isin(deusNaoPermite)), 'MinutosTotais'] *= 2

        #Somando os minutos extras
        resultado = horas.groupby('Semana')['MinutosTotais'].sum().reset_index()
        resultado = resultado['MinutosTotais'].sum()

        #Exportando em excel
        horas = horas.rename(columns={'Data':'Data','diaSemana':'Dia da Semana','Total':'Horas Trabalhadas','MinutosTotais':'Minutos Extra'})
        horas['Data'] = horas['Data'].dt.strftime('%d/%m/%Y')
        try:
            horas['Horas Trabalhadas'] = pd.to_timedelta(horas['Horas Trabalhadas']).dt.components['hours'].astype(str) + ':' + pd.to_timedelta(horas['Horas Trabalhadas']).dt.components['minutes'].astype(str).str.zfill(2)
        except:
        #Tratamento de erro do jeito que deus gosta
            ...
        horas['Dia da Semana'] = horas['Dia da Semana'].replace({'1': 'Segunda', '2': 'Terça', '3': 'Quarta', '4': 'Quinta', '5': 'Sexta', '6': 'Sábado', '7': 'Domingo'})
        horas.to_excel(f'bancoHoras_{aba}.xlsx', index=False)
        print(f'{aba}: {resultado // 60} horas e {resultado % 60} minutos')


    except:
        print(f'Erro ao processar {aba}')

print(f'Total de minutos extras de todas as abas: {resultado // 60} horas e {resultado % 60} minutos')
