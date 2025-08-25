import streamlit as st
import pandas as pd
import datetime as dt
from collections import defaultdict
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

# Adicione no início do código
import base64

# Substitua as leituras de arquivo por:
@st.cache_data
def load_salas(uploaded_file):
    return pd.read_excel(uploaded_file)

@st.cache_data
def load_turmas(uploaded_file):
    return pd.read_excel(uploaded_file)

# ================= INTERFACE =================
st.set_page_config(page_title="Alocação de Turmas", layout="wide")
st.title("📅 Sistema de Alocação de Turmas em Salas")

file_salas = st.file_uploader("📂 Envie o arquivo de SALAS", type=["xlsx"])
file_turmas = st.file_uploader("📂 Envie o arquivo de TURMAS", type=["xlsx"])

# Inicializar variáveis de sessão
if "resultados" not in st.session_state:
    st.session_state.resultados = None
if "buffers_salas" not in st.session_state:
    st.session_state.buffers_salas = {}

# ================= PROCESSAMENTO =================
if file_salas and file_turmas:
    if st.button("🚀 Rodar Alocação"):
        # ================= LEITURA =================
        df_salas = load_salas(file_salas)
        df_turmas = load_turmas(file_turmas)

        salas_data = df_salas["SALAS"].to_numpy()
        capacidade_data = df_salas["CAPACIDADE"].to_numpy()
        cod_data = df_turmas["CÓDIGO"].to_numpy()
        cod_turma_data = df_turmas["Nº DA TURMA"].to_numpy()
        turmas_data = df_turmas["DISCIPLINA"].to_numpy()
        demanda_data = df_turmas["PREVISÃO DE ALUNOS"].to_numpy()
        professor_data = df_turmas["PROFESSOR"].to_numpy()
        dias_data = df_turmas["DIAS_PADRONIZADOS"].to_numpy()
        horarios_data = df_turmas["HORÁRIOS"].to_numpy()

        # Montar horários por turma
        horarios_turmas = []
        for i in range(len(dias_data)):
            dias_list = str(dias_data[i]).split()
            horarios_list = str(horarios_data[i]).split(', ')
            turma_horarios = []
            for dia in dias_list:
                for hora in horarios_list:
                    turma_horarios.append(f'{dia} {hora}')
            horarios_turmas.append(turma_horarios)

        # Criar lista de salas
        salas_ct = []
        for i in range(len(salas_data)):
            salas_ct.append({
                "NOME": salas_data[i],
                "CAPACIDADE": capacidade_data[i],
                "HORARIOS_OCUPADOS": set()
            })

        # Criar lista de disciplinas
        disciplinas = []
        for i in range(len(turmas_data)):
            disciplinas.append({
                "CODIGO": cod_data[i],
                "DISCIPLINA": turmas_data[i],
                "CODIGO TURMA": cod_turma_data[i],
                "HORARIOS": horarios_turmas[i],
                "ALUNOS": demanda_data[i],
                "PROFESSOR": professor_data[i]
            })

        disciplinas.sort(key=lambda d: d["ALUNOS"], reverse=True)

        # ================= ALOCAÇÃO =================
        alocacao = []
        for disc in disciplinas:
            alunos = disc["ALUNOS"]
            horarios_disciplina = disc["HORARIOS"]
            melhor_sala = None
            menor_ociosidade = float("inf")

            for sala in salas_ct:
                if sala["CAPACIDADE"] >= alunos:
                    is_disponivel = all(h not in sala["HORARIOS_OCUPADOS"] for h in horarios_disciplina)
                    if is_disponivel:
                        ociosidade = sala["CAPACIDADE"] - alunos
                        if ociosidade < menor_ociosidade:
                            menor_ociosidade = ociosidade
                            melhor_sala = sala

            if melhor_sala:
                alocacao.append({
                    "CODIGO": disc["CODIGO"],
                    "DISCIPLINA": disc["DISCIPLINA"],
                    "SALA": melhor_sala["NOME"],
                    "CODIGO TURMA": disc["CODIGO TURMA"],
                    "PROFESSOR": disc["PROFESSOR"],
                    "HORARIO": ", ".join(horarios_disciplina),
                    "ALUNOS": alunos,
                    "OCIOSIDADE": menor_ociosidade,
                    "STATUS": "Alocada"
                })
                for h in horarios_disciplina:
                    melhor_sala["HORARIOS_OCUPADOS"].add(h)
            else:
                alocacao.append({
                    "CODIGO": disc["CODIGO"],
                    "DISCIPLINA": disc["DISCIPLINA"],
                    "SALA": None,
                    "CODIGO TURMA": disc["CODIGO TURMA"],
                    "PROFESSOR": disc["PROFESSOR"],
                    "HORARIO": ", ".join(horarios_disciplina),
                    "ALUNOS": alunos,
                    "OCIOSIDADE": None,
                    "STATUS": "Não alocada"
                })

        # ================= RESULTADOS GERAIS =================
        df_resultados = pd.DataFrame(alocacao)
        buffer_geral = BytesIO()
        df_resultados.to_excel(buffer_geral, index=False)
        buffer_geral.seek(0)

        # Salva no session_state
        st.session_state.resultados = buffer_geral

        # ================= RESULTADOS POR SALA =================
        dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"]
        horas_minutos = []
        for h in range(7, 22):
            horas_minutos.append(f'{h:02d}:00 - {h:02d}:30')
            horas_minutos.append(f'{h:02d}:30 - {h+1:02d}:00')

        def split_horario(horario_completo):
            partes = horario_completo.split()
            dia = partes[0]
            hora_str = partes[1]
            if '-' not in hora_str:
                return []
            hora_inicio_str, hora_fim_str = hora_str.split('-')
            hora_inicio = dt.datetime.strptime(hora_inicio_str, '%H:%M:%S')
            intervalos = []
            intervalos.append(f'{dia} {hora_inicio.strftime("%H:%M")} - {hora_inicio.replace(minute=30).strftime("%H:%M")}')
            hora_segundo_intervalo = hora_inicio.replace(minute=30)
            intervalos.append(f'{dia} {hora_segundo_intervalo.strftime("%H:%M")} - {dt.datetime.strptime(hora_fim_str, "%H:%M:%S").strftime("%H:%M")}')
            return intervalos

        horarios_por_sala = defaultdict(lambda: defaultdict(dict))
        for aloc in alocacao:
            if aloc['SALA']:
                sala_nome = aloc['SALA']
                disciplina_info = f"{aloc['CODIGO']} - {aloc['DISCIPLINA']} - {aloc['CODIGO TURMA']} - {aloc['PROFESSOR']}"
                horarios_blocos = [h.strip() for h in aloc['HORARIO'].split(',')]
                for bloco in horarios_blocos:
                    if bloco:
                        dia = bloco.split()[0]
                        horarios_30min = split_horario(bloco)
                        for horario_30min in horarios_30min:
                            _, horario_formatado = horario_30min.split(' ', 1)
                            horarios_por_sala[sala_nome][dia][horario_formatado] = disciplina_info

        borda_fina = Border(left=Side(style='thin'), right=Side(style='thin'),
                            top=Side(style='thin'), bottom=Side(style='thin'))
        alinhamento_centro = Alignment(horizontal='center', vertical='center', wrap_text=True)
        fonte_padrao = Font(size=10)

        buffers_salas = {}
        for sala in salas_ct:
            sala_nome = sala["NOME"]
            wb = Workbook()
            ws = wb.active
            ws.title = "Horário"

            info_sala = f"Centro de Tecnologia | {sala_nome} | Capacidade: {sala['CAPACIDADE']}"
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(dias_semana) + 1)
            cell_info = ws.cell(row=1, column=1, value=info_sala)
            cell_info.font = Font(bold=True, size=12)
            cell_info.alignment = Alignment(horizontal='center', vertical='center')

            ws.cell(row=2, column=1, value="Horário").font = Font(bold=True)
            for col, dia in enumerate(dias_semana, start=2):
                ws.cell(row=2, column=col, value=dia).font = Font(bold=True)

            for row, hora in enumerate(horas_minutos, start=2):
                ws.cell(row=row, column=1, value=hora)

            if sala_nome in horarios_por_sala:
                for dia, horarios in horarios_por_sala[sala_nome].items():
                    if dia == 'SEGUNDA': col = 2
                    elif dia == 'TERÇA': col = 3
                    elif dia == 'QUARTA': col = 4
                    elif dia == 'QUINTA': col = 5
                    elif dia == 'SEXTA': col = 6
                    elif dia == 'SÁBADO': col = 7
                    else: continue
                    for horario, info in horarios.items():
                        if horario in horas_minutos:
                            row_idx = horas_minutos.index(horario) + 2
                            ws.cell(row=row_idx, column=col, value=info)

            # Mesclar células
            for col in range(2, len(dias_semana) + 2):
                start_row = 2
                current_value = ws.cell(row=2, column=col).value
                for row in range(3, len(horas_minutos) + 2):
                    value = ws.cell(row=row, column=col).value
                    if value != current_value:
                        if current_value not in (None, "") and row - 1 > start_row:
                            ws.merge_cells(start_row=start_row, start_column=col,
                                           end_row=row - 1, end_column=col)
                        start_row = row
                        current_value = value
                if current_value not in (None, "") and len(horas_minutos) + 1 > start_row:
                    ws.merge_cells(start_row=start_row, start_column=col,
                                   end_row=len(horas_minutos) + 1, end_column=col)

            # Estilo
            for row in ws.iter_rows(min_row=1, max_row=len(horas_minutos) + 1,
                                    min_col=1, max_col=len(dias_semana) + 1):
                for cell in row:
                    cell.border = borda_fina
                    cell.alignment = alinhamento_centro
                    cell.font = fonte_padrao

            for i, col_cells in enumerate(ws.columns, start=1):
                max_length = 0
                col_letter = get_column_letter(i)
                for cell in col_cells:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = max_length + 2

            # Salvar em memória
            buffer_sala = BytesIO()
            wb.save(buffer_sala)
            buffer_sala.seek(0)

            buffers_salas[sala_nome] = buffer_sala

        # Salva todos os buffers no session_state
        st.session_state.buffers_salas = buffers_salas

# ================= DOWNLOADS =================
if st.session_state.resultados:
    st.download_button(
        label="⬇️ Baixar Resultados Gerais",
        data=st.session_state.resultados,
        file_name="Resultados_Gerais.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    for sala_nome, buffer_sala in st.session_state.buffers_salas.items():
        st.download_button(
            label=f"⬇️ Baixar {sala_nome}",
            data=buffer_sala,
            file_name=f"horarios_{sala_nome.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
