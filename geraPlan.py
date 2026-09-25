import pandas as pd
import re
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import date, datetime
from flask import Flask, request, send_file, jsonify
import os

class Aluno:
    def __init__(self,nome,matricula,curso,cpf,banco,ag,conta):
        self.nome = nome
        self.matricula = matricula
        self.curso = curso
        self.cpf = cpf
        self.banco = banco
        self.ag = ag
        self.conta = conta

    def __str__(self):
        return f"{self.nome} ({self.matricula} - {self.cursp})"

# leio a planilha
def ler_alunos(caminho_planilha, nome_turma):
    """Lê uma planilha e retorna uma lista de objetos Aluno."""
    df = pd.read_excel(caminho_planilha, engine='openpyxl', sheet_name='prim')
    alunos = []

    for _, row in df.iterrows():
        nome = row["Aluno"]
        matricula = row["Matrícula"]
        turma = nome_turma # if nome_turma else row.get('Turma', 'Turma Desconhecida')  # Se não houver coluna "Turma", usa o valor passado

        alunos.append(Aluno(nome, matricula, turma))
    return alunos


def mapa_turma(arq, etapa):
    ordem_etapas = ["Etapa 1", "Etapa 2", "Etapa 3", "Etapa 4", "Etapa Final"]

    # Descobrimos quais abas precisamos somar as faltas
    if etapa in ordem_etapas:
        idx_etapa = ordem_etapas.index(etapa)
        etapas_soma = ordem_etapas[:idx_etapa + 1] # Pega da Etapa 1 até a etapa atual
    else:
        etapas_soma = [etapa] # Se for um nome desconhecido, lê só ela mesma

    excel = pd.ExcelFile(arq)
    faltas_acumuladas = {}
    # NEW
    notas_por_aluno = {} # Guarda a lista de etapas de cada aluno: { 'Aluno': [ [notas_e0], [notas_e1] ]

    # Percorre todas as abas necessárias para somar as faltas
    for idx_etapa_num, nome_aba in enumerate(etapas_soma):
        if nome_aba in excel.sheet_names:
            df_aba = pd.read_excel(excel, sheet_name=nome_aba, header=2)
            colunas_faltas = [col for col in df_aba.columns if str(col).startswith("F")]
            colunas_notas = [col for col in df_aba.columns if str(col).startswith("N")]

            for _, row in df_aba.iterrows():
                nome_aluno = row["Aluno"]
                if nome_aluno not in faltas_acumuladas:
                    faltas_acumuladas[nome_aluno] = 0
                    notas_por_aluno[nome_aluno] = []

                # Soma as faltas dessa aba
                for col_falta in colunas_faltas:
                    valor_falta = pd.to_numeric(row[col_falta], errors='coerce')
                    if pd.notna(valor_falta):
                        faltas_acumuladas[nome_aluno] += valor_falta

                notas_etapa = [row[col] for col in colunas_notas]
                notas_por_aluno[nome_aluno].append(notas_etapa)

    # colunas_faltas = [col for col in df.columns if col.startswith("F")]

    # Monta a lista final de retorno
    alunos_dados = []
    for nome_aluno, lista_de_etapas in notas_por_aluno.items():
        s_f = faltas_acumuladas.get(nome_aluno, 0)
        # Retorna: (Nome, [[notas_e0], [notas_e1], ...], Faltas_Acumuladas)
        alunos_dados.append((nome_aluno, lista_de_etapas, s_f))

    return alunos_dados


app = Flask(__name__)
# 1. Rota para mostrar o formulário HTML que criámos
@app.route('/')
def mostrar_formulario():
    # Assume que o index.html está na mesma pasta que este script Python
    return send_file('index.html')

@app.route('/geraPlan.py', methods=['POST'])
def processar_formulario():
    try:
        # 1. Validação dos inputs do formulário

        ficheiro_mapa = request.files.get('arquivo_mapa')
        if not ficheiro_mapa or ficheiro_mapa.filename == '':
            return jsonify({"erro": "Erro: Nenhum arquivo foi enviado."}), 400


        # 2. Mapeamento de colunas equivalentes e disciplinas
        infodisciplinas = []
        mapav = []
        novo_idx = 0
        Sh = 0
        cargas = filtro_ch(caminho_temp, etapa_escolhida)

        for i, sigla in enumerate(disc):
            obj = dicion.get(sigla)
            nome_disciplina = obj.nome if obj else sigla
            if len(infodisciplinas) == 0 or nome_disciplina != infodisciplinas[-1]['nome']:
                infodisciplinas.append({
                    'nome': nome_disciplina,
                    'area': obj.area if obj else "Outras Áreas",
                    'area_id': obj.area_id if obj else "geral",
                    'idx': novo_idx
                })
                val_cel = str(cargas[i]).strip()
                if "H" in val_cel:
                    x_str = val_cel.split("H")[0]
                    try:
                        Sh += int(x_str)
                    except ValueError:
                        pass
                mapav.append([i])
                novo_idx += 1
            else:
                mapav[-1].append(i)

        infodisciplinas.sort(key=lambda x: x['area'])

        dados_originais = mapa_turma(caminho_temp, etapa_escolhida)
        dados = []

        # 3. Processamento dos Alunos e Acúmulo de Notas
        for nome_aluno, lista_etapas_brutas, faltas in dados_originais:
            # `lista_etapas_brutas` contém [[notas_e0], [notas_e1], ...]
            # `notas_acumuladas` conterá [[notas_limpas_e0], [notas_limpas_e1], ...]
            notas_acumuladas = []

            # Acumuladores de notas das Áreas (somam todas as etapas disponíveis 0..N)
            soma_areas = {'matematica': 0, 'linguagens': 0, 'humanas': 0, 'natureza': 0, 'tecnicas': 0}
            cont_areas = {'matematica': 0, 'linguagens': 0, 'humanas': 0, 'natureza': 0, 'tecnicas': 0}

            # Itera sobre cada etapa usando o índice numérico (0 para Etapa 1, 1 para Etapa 2, etc)
            for idx_num, notas_originais_etapa in enumerate(lista_etapas_brutas):
                notas_limpas_etapa = []

                # Limpeza de notas para a etapa atual
                for vizinhos in mapav:
                    nota_final = "-"
                    for col in vizinhos:
                        if col < len(notas_originais_etapa):
                            nota = notas_originais_etapa[col]
                            if pd.isna(nota):
                                continue
                            nota_str = str(nota).strip().replace(',', '.')
                            if nota_str in ["", "-", "nan", "None"]:
                                continue
                            try:
                                float(nota_str)
                                nota_final = nota
                                break
                            except ValueError:
                                pass
                    notas_limpas_etapa.append(nota_final)

                # Adiciona a lista de notas limpas desta etapa na posição correspondente
                notas_acumuladas.append(notas_limpas_etapa)

                # Computa na soma das Áreas
                for info_disc in infodisciplinas:
                    nota = notas_limpas_etapa[info_disc['idx']]
                    area_id = info_disc['area_id']

                    if nota != "-" and area_id in soma_areas:
                        soma_areas[area_id] += float(nota)
                        cont_areas[area_id] += 1

            # 4. Médias Finais por Área de Conhecimento (Acumuladas)
            medias_areas = {}
            for area_id in soma_areas.keys():
                if cont_areas[area_id] > 0:
                    media_calc = soma_areas[area_id] / cont_areas[area_id]
                    medias_areas[area_id] = round(media_calc, 2)
                else:
                    medias_areas[area_id] = "-"

            # Estrutura do aluno enviada ao Jinja2:
            # (nome_aluno, lista_notas_por_indice_0_a_3, faltas, Fr, medias_areas)
            dados.append((nome_aluno, notas_acumuladas, faltas, Fr, medias_areas))

        add_aluno_page(dados, infodisciplinas, Sh, nome_arquivo=nome_pdf)

        # ~ return jsonify({
            # ~ "mensagem": f"Sucesso! Boletins gerados em {fim - inicio:.2f} segundos.",
            # ~ "arquivo": nome_pdf
        # ~ }), 200

    except Exception as e:
        return jsonify({"erro": str(e)}), 500

@app.route('/download/<nome_arquivo>')
def baixar_pdf(nome_arquivo):
    # O 'as_attachment=True' força o navegador a fazer o download do arquivo
    return send_file(nome_arquivo, as_attachment=True)

# Arranca o servidor
if __name__ == "__main__":
    print("Servidor a iniciar! Abra o seu navegador e vá a: http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
