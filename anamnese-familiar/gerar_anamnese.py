#!/usr/bin/env python3
"""Gera a Anamnese Social da Família (DOCX + PDF) da ONG Santidade Integral em Cristo.

Uso:  python3 gerar_anamnese.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from timbrado import *  # noqa: E402,F401

HERE = Path(__file__).resolve().parent
OUT = HERE / "Anamnese_Social_Familiar_Santidade_Integral.docx"


# ----------------------------------------------------------------- documento
def build():
    doc = new_document(
        header_meta="Prontuário nº: ______________    Unidade / Frente de trabalho: ______________________",
        footer_label="Documento de uso interno — informações protegidas pela LGPD (Lei 13.709/2018). "
                     "Anamnese Social Familiar  •  v1.0",
    )

    # ================================================== título
    title_block(doc, "ANAMNESE SOCIAL DA FAMÍLIA",
                "Cadastro, Diagnóstico Socioeconômico e Plano de Desenvolvimento Familiar",
                "Eu sou o pão da vida; aquele que vem a mim não terá fome, "
                "e quem crê em mim nunca terá sede.", "João 6:35")

    # ================================================== 1. ATENDIMENTO
    section_title(doc, "1", "Dados do atendimento")
    field_line(doc, ("Data da anamnese", 1), ("Voluntário(a) responsável", 3), ("Contato", 1.3))
    check_line(doc, "Forma de acesso", ["Busca ativa", "Demanda espontânea", "Encaminhamento (CRAS/UBS/Igreja): ______________"])
    check_line(doc, "Assistência em curso", ["Cesta básica", "Medicamentos", "Ambos", "Nenhuma (novo cadastro)"])
    field_line(doc, ("Desde (mês/ano)", 1), ("Periodicidade da entrega", 2))
    check_line(doc, "Tipo de atendimento", ["Inclusão (1º cadastro)", "Atualização cadastral", "Reavaliação do plano"])

    # ================================================== 2. RESPONSÁVEL FAMILIAR
    section_title(doc, "2", "Identificação do(a) Responsável Familiar (RF)")
    field_line(doc, ("Nome completo", 1))
    field_line(doc, ("Data de nascimento", 1), ("CPF", 1.3), ("RG / Órgão emissor", 1.4), ("NIS (CadÚnico)", 1.1))
    check_line(doc, "Sexo", ["Feminino", "Masculino"], note="   ")
    p = doc.paragraphs[-1]
    run(p, "Estado civil: ", bold=True)
    for o in ("Solteiro(a)", "Casado(a)/União estável", "Separado(a)", "Viúvo(a)"):
        run(p, f"{BOX} {o}  ")
    check_line(doc, "Raça/cor (autodeclarada)", ["Branca", "Preta", "Parda", "Amarela", "Indígena"])
    check_line(doc, "Escolaridade", ["Não alfabetizado(a)", "Fund. incompleto", "Fund. completo",
                                     "Médio incompleto", "Médio completo", "Superior"])
    field_line(doc, ("Telefone / WhatsApp", 1), ("Telefone de recado (nome)", 1.6))
    field_line(doc, ("Endereço", 3), ("Nº", 0.3), ("Bairro/Comunidade", 1.2))
    field_line(doc, ("Município/UF", 1.2), ("CEP", 0.7), ("Ponto de referência", 2))
    check_line(doc, "Zona", ["Urbana", "Rural"], note="      ")
    p = doc.paragraphs[-1]
    run(p, "Tempo de residência no local: ", bold=True, size=9)
    run(p, "____________", color=GREY, size=9)
    run(p, "      CRAS de referência: ", bold=True, size=9)
    run(p, "______________________", color=GREY, size=9)

    # ================================================== 3. COMPOSIÇÃO FAMILIAR
    section_title(doc, "3", "Composição familiar (todos os moradores do domicílio)")
    table(doc,
          ["Nome", "Parentesco c/ RF", "Idade", "Estuda?", "Escolaridade", "Ocupação / fonte de renda",
           "Renda mensal (R$)", "Condição de saúde / deficiência"],
          [], [4.0, 2.2, 1.1, 1.3, 2.0, 3.0, 1.7, 2.1], blank_rows=7)
    field_line(doc, ("Total no domicílio", 1), ("Crianças 0–6", 1), ("7–17 anos", 1),
               ("Idosos 60+", 1), ("Gestantes/nutrizes", 1), ("PcD", 1))
    check_line(doc, "Criança/adolescente fora da escola", ["Não", "Sim — quem/motivo: ________________________"])

    # ================================================== 4. HABITAÇÃO
    section_title(doc, "4", "Condições de habitação e acesso a serviços")
    check_line(doc, "Situação do imóvel", ["Próprio quitado", "Próprio financiado", "Alugado (R$ ______)", "Cedido", "Ocupação", "Outro"])
    check_line(doc, "Tipo / material", ["Alvenaria", "Madeira", "Taipa/pau a pique", "Material reaproveitado", "Misto"])
    field_line(doc, ("Cômodos", 1), ("Dormitórios", 1), ("Pessoas por dormitório", 1),
               f"Banheiro:  {BOX} Interno  {BOX} Externo  {BOX} Não possui")
    check_line(doc, "Água", ["Rede pública", "Poço/nascente", "Cisterna", "Carro-pipa", "Sem acesso regular"])
    check_line(doc, "Esgoto", ["Rede", "Fossa séptica", "Fossa rudimentar", "Céu aberto"], note="    ")
    p = doc.paragraphs[-1]
    run(p, "Energia: ", bold=True)
    for o in ("Regular", "Irregular", "Sem energia"):
        run(p, f"{BOX} {o}  ")
    check_line(doc, "Coleta de lixo", ["Regular", "Queima/enterra", "Descarte irregular"], note="      ")
    p = doc.paragraphs[-1]
    run(p, "Área de risco (enchente/deslizamento): ", bold=True, size=9)
    run(p, f"{BOX} Não  {BOX} Sim", size=9)
    check_line(doc, "Bens e equipamentos", ["Fogão", "Geladeira", "Botijão de gás", "Máq. de lavar", "Celular c/ internet",
                                           "TV", "Veículo/moto"])

    # ================================================== 5. SAÚDE
    section_title(doc, "5", "Saúde e uso de medicamentos")
    check_line(doc, "Família cadastrada na UBS/ESF", ["Sim", "Não"], note="   ")
    p = doc.paragraphs[-1]
    run(p, "Agente Comunitário de Saúde: ", bold=True)
    run(p, "______________________________", color=GREY)
    check_line(doc, "Condições presentes no domicílio",
               ["Hipertensão", "Diabetes", "Doença respiratória", "Transtorno mental", "Deficiência",
                "Gestação", "Acamado(a)", "Uso abusivo de álcool/drogas", "Outra: ________"])
    table(doc,
          ["Medicamento (nome e dosagem)", "Quem utiliza", "Posologia", "Receita médica?", "Forma de obtenção *", "Custo mensal (R$)"],
          [], [5.0, 3.0, 2.4, 2.0, 3.0, 2.0], blank_rows=5)
    note(doc, "* Forma de obtenção: SUS / Farmácia Popular / Compra própria / Doação da ONG / Sem acesso (interrupção do tratamento).")

    # ================================================== 6. SEGURANÇA ALIMENTAR
    section_title(doc, "6", "Segurança alimentar e nutricional")
    note(doc, "Perguntas de triagem adaptadas da Escala Brasileira de Insegurança Alimentar (EBIA) — referência: últimos 3 meses.")
    table(doc,
          ["Nos últimos 3 meses…", "Sim", "Não"],
          [["Os moradores tiveram preocupação de que os alimentos acabassem antes de poderem comprar mais?", BOX, BOX],
           ["Os alimentos acabaram antes que houvesse dinheiro para comprar mais?", BOX, BOX],
           ["Algum adulto diminuiu a quantidade de alimentos ou pulou refeições por falta de dinheiro?", BOX, BOX],
           ["Algum adulto deixou de fazer uma refeição o dia inteiro por falta de dinheiro?", BOX, BOX],
           ["Alguma criança/adolescente comeu menos ou deixou de comer por falta de dinheiro?", BOX, BOX]],
          [14.9, 1.25, 1.25])
    field_line(doc, ("Refeições por dia", 0.4), ("Principal fonte dos alimentos", 2.2))
    check_line(doc, "Produção para consumo próprio (horta, criação de animais)", ["Não", "Sim — o quê: ______________________"])
    note(doc, "Classificação (soma de “Sim”): 0 = segurança alimentar • 1–2 = IA leve • 3–4 = IA moderada • 5 = IA grave.")

    # ================================================== 7. TRABALHO E RENDA
    section_title(doc, "7", "Trabalho, renda e orçamento familiar")
    check_line(doc, "Ocupação do RF", ["Desempregado", "Informal/bicos", "Autônomo", "CLT", "Produtor rural",
                                       "Aposentado/pensionista", "Do lar"])
    field_line(doc, ("Tempo sem trabalho formal", 1), ("Última ocupação", 2.5))
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fix_widths(t, [8.7, 8.7])
    cant_split(t.rows[0])
    left, right = t.rows[0].cells
    # Receitas
    p = left.paragraphs[0]
    spacing(p, after=1)
    run(p, "RECEITAS MENSAIS", bold=True, size=9, color=GREEN)
    receitas = [("Renda do trabalho (soma dos membros)", ""), ("Bolsa Família", ""), ("BPC/LOAS", ""),
                ("Aposentadoria / pensão", ""), ("Seguro-desemprego / auxílio", ""), ("Pensão alimentícia", ""),
                ("Ajuda de terceiros / igreja", ""), ("Outras", ""), (("TOTAL DE RECEITAS", True), "")]
    despesas = [("Aluguel / prestação", ""), ("Água e energia", ""), ("Gás", ""), ("Alimentação", ""),
                ("Medicamentos / saúde", ""), ("Transporte", ""), ("Educação / material escolar", ""),
                ("Dívidas / parcelas", ""), (("TOTAL DE DESPESAS", True), "")]
    for cell, items, title in ((left, receitas, None), (right, despesas, "DESPESAS MENSAIS")):
        if title:
            p = cell.paragraphs[0]
            spacing(p, after=1)
            run(p, title, bold=True, size=9, color=GREEN)
        inner = cell.add_table(rows=0, cols=2)
        inner.autofit = False
        borders(inner)
        inner_widths = [6.0, 2.2]
        set_cell_margins(inner, top=20, bottom=20)
        for label, _ in items:
            r = inner.add_row().cells
            r[0].width, r[1].width = Cm(6.2), Cm(2.3)
            bold = isinstance(label, tuple)
            txt = label[0] if bold else label
            if bold:
                shade(r[0], LIGHT_HEX)
                shade(r[1], LIGHT_HEX)
            p = r[0].paragraphs[0]
            spacing(p, after=0)
            run(p, txt, bold=bold, size=8.5)
            p = r[1].paragraphs[0]
            spacing(p, after=0)
            run(p, "R$", size=8.5, color=GREY)
        fix_widths(inner, inner_widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    field_line(doc, ("Renda per capita R$", 1), ("Saldo mensal R$", 1), ("Dívidas acumuladas R$", 1))
    note(doc, "Renda per capita = total de receitas ÷ nº de pessoas no domicílio.  Saldo mensal = receitas − despesas.")
    check_line(doc, "Faixa de renda per capita", ["Até R$ 218 (extrema pobreza)", "R$ 218 a R$ 706 (até ½ salário mínimo)", "Acima de ½ SM"])
    check_line(doc, "Cadastro Único", ["Não cadastrado", "Cadastrado e atualizado", "Desatualizado (> 24 meses)"])

    # ================================================== 8. VÍNCULOS E REDE
    section_title(doc, "8", "Vínculos familiares, comunitários e rede de proteção")
    check_line(doc, "Acompanhamento em serviços", ["CRAS/PAIF", "CREAS/PAEFI", "CAPS", "Conselho Tutelar", "Nenhum"])
    check_line(doc, "Participação comunitária", ["Igreja/comunidade de fé", "Associação de moradores", "Grupo produtivo/cooperativa", "Nenhuma"])
    check_line(doc, "Situações de risco", ["Violência doméstica", "Trabalho infantil", "Negligência", "Isolamento",
                                           "Conflito familiar", "Nenhuma"])
    lines_box(doc, 3, "Rede de apoio (parentes, vizinhos, lideranças) e observações sobre a dinâmica familiar:")

    # ================================================== 9. POTENCIALIDADES
    section_title(doc, "9", "Potencialidades, ativos e aspirações da família")
    note(doc, "Levantamento dos ativos humanos, sociais, físicos e produtivos — base para o projeto de geração de renda.")
    table(doc,
          ["Membro", "Habilidades / ofícios", "Cursos / qualificação", "Experiência profissional (anos / área)", "Disponibilidade"],
          [], [3.0, 5.4, 3.0, 3.4, 2.6], blank_rows=4)
    note(doc, "Exemplos de ofícios: costura, pedreiro, eletricista, cozinha/salgados, cabeleireiro, manicure, roça, criação de animais, motorista, cuidador.")
    check_line(doc, "Ativos produtivos", ["Quintal (____ m²)", "Terra p/ cultivo (___ ha)", "Animais: ______",
                                          "Ferramentas", "Máq. costura", "Veículo"])
    check_line(doc, "Documentação e crédito", ["CNH", "CNPJ/MEI ativo", "DAP/CAF", "Conta bancária",
                                               "Nome negativado", "Sem documentos"])
    check_line(doc, "Interesse da família", ["Emprego formal", "Empreendimento próprio", "Produção agropecuária",
                                             "Qualificação", "Retomar estudos"])
    lines_box(doc, 3, "O que a família sonha para os próximos 12 meses? Quais obstáculos ela identifica?")

    # ================================================== 10. DIAGNÓSTICO
    section_title(doc, "10", "Diagnóstico social (síntese do voluntário)")
    check_line(doc, "Classificação da vulnerabilidade", ["Emergencial (risco imediato)", "Transitória (perda de renda)",
                                                          "Crônica (intergeracional)"])
    check_line(doc, "Prioridade de intervenção", ["Alta", "Média", "Baixa"])
    check_line(doc, "Encaminhamentos", ["CadÚnico", "CRAS", "UBS", "Escola", "Defensoria", "SINE/emprego", "EMATER/Sind. rural"])
    lines_box(doc, 6, "Principais vulnerabilidades, riscos e potencialidades observadas:")

    # ================================================== 11. PLANO
    doc.add_page_break()
    section_title(doc, "11", "Plano de Desenvolvimento Familiar (pactuado com a família)")
    check_line(doc, "Eixo do projeto", ["Trabalho/empregabilidade", "Empreendedorismo", "Agropecuário", "Qualificação", "Misto"])
    field_line(doc, ("Descrição do projeto de geração de renda", 1))
    lines_box(doc, 1)
    field_line(doc, ("Investimento inicial (R$)", 1), f"Fonte:  {BOX} ONG  {BOX} Família  {BOX} Crédito  {BOX} Parceiro",
               ("Renda mensal projetada (R$)", 1))
    table(doc,
          ["Meta / objetivo", "Ações concretas", "Responsável", "Prazo", "Indicador de sucesso"],
          [], [4.2, 5.4, 2.4, 1.8, 3.6], blank_rows=5)
    field_line(doc, ("Redução gradual da cesta a partir de", 1), ("Meta de autonomia (data)", 1),
               ("Reavaliação a cada", 0.4), "meses")

    # ================================================== 12. ACOMPANHAMENTO
    section_title(doc, "12", "Registro de acompanhamento (evolução)")
    table(doc,
          ["Data", "Tipo *", "Situação encontrada / evolução das metas", "Próxima ação", "Voluntário"],
          [], [1.7, 2.2, 8.1, 3.2, 2.2], blank_rows=6)
    note(doc, "* Tipo: V = visita domiciliar • C = contato telefônico • E = entrega (cesta/medicamento) • A = atendimento na sede.")

    # ================================================== 13. CONSENTIMENTO
    section_title(doc, "13", "Termo de consentimento e assinaturas")
    p = doc.add_paragraph()
    spacing(p, after=4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p, "Declaro que as informações prestadas são verdadeiras e autorizo a ONG Santidade Integral em Cristo a registrar, "
           "guardar e utilizar estes dados exclusivamente para fins de assistência social, acompanhamento familiar e "
           "encaminhamento à rede de proteção, nos termos da Lei Geral de Proteção de Dados (Lei nº 13.709/2018). "
           "Estou ciente de que posso solicitar acesso, correção ou exclusão dos meus dados a qualquer momento.", size=9)
    field_line(doc, ("Local e data", 1))
    t = doc.add_table(rows=1, cols=2)
    t.autofit = False
    fix_widths(t, [8.7, 8.7])
    for i, label in enumerate(("Responsável Familiar", "Voluntário(a) responsável pela anamnese")):
        c = t.rows[0].cells[i]
        c.width = Cm(8.7)
        p = c.paragraphs[0]
        spacing(p, before=18, after=0)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run(p, "_" * 42, color=GREY)
        p = c.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        spacing(p, after=0)
        run(p, label, size=8.5, color=GREY)

    doc.save(OUT)
    print(f"OK -> {OUT}")
    print(f"OK -> {to_pdf(OUT)}")


if __name__ == "__main__":
    build()
