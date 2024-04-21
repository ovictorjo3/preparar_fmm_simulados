import openpyxl

def conferir_gabarito(gabarito, respostas_aluno):
    pontuacao = 0
    for i in range(len(gabarito)):
        if gabarito[i] == respostas_aluno[i]:
            pontuacao += 1
    return pontuacao

def salvar_resultados(resultados):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Resultados" #nome da turma

    # Escrevendo cabeçalho
    sheet["A1"] = "Aluno(a)"
    sheet["B1"] = "Pontuação"

    # Escrevendo dados
    for i, resultado in enumerate(resultados, start=2):
        sheet[f"A{i}"] = resultado[0]
        sheet[f"B{i}"] = resultado[1]

    # Salvando o arquivo
    workbook.save(filename="resultados.xlsx")

# Gabarito
gabarito = ['D','E','D','D','A','A','B','C','A','D',
            'E','B','C','D','E','A','B','C','D','B',
            'B','E','C','E','C','C','A','A','D','E',
            'C','B','C','B','D','E','E','D','D','A']

# Respostas dos alunos (formato: (nome, respostas))
respostas_alunos = [
    ("GABRIELE PINTO FERREIRA",['D','B','B','A','C','A','','C','D','',
                    'C','B','B','C','A','A','C','A','A','B',
                    'C','A','D','A','C','D','D','C','D','D',
                    'B','A','D','A','D','D','B','D','E','B']),

    ("CLARICE CONCEIÇÃO",['C','E','B','C','D','C','B','D','B','D',
                    'A','D','A','B','B','A','D','E','D','B',
                    'B','C','E','E','C','A','B','D','B','D',
                    'E','B','C','B','B','C','B','D','E','B']), #AQUI

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("AJOAO VICTOR",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("PEDRO AUGUSTO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("AJOAO VICTOR",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("PEDRO AUGUSTO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("AJOAO VICTOR",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("PEDRO AUGUSTO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),

    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    ("JOSE OTAVIO",['D','E','D','D','A','A','B','C','A','D',
                    'E','B','C','D','E','A','B','C','D','B',
                    'B','E','C','E','C','C','A','A','D','E',
                    'C','B','C','B','D','E','E','D','D','A']),
    
    
    
    
]

resultados = []

# Conferindo respostas e calculando pontuações
for aluno, respostas in respostas_alunos:
    pontuacao = conferir_gabarito(gabarito, respostas)
    resultados.append((aluno, pontuacao))

# Salvando resultados em uma planilha Excel
salvar_resultados(resultados)

print("Resultados salvos com sucesso!")