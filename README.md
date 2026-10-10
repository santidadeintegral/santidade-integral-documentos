# Santidade Integral em Cristo — Documentos

Repositório de documentos institucionais da ONG **Santidade Integral em Cristo**
(ação social: entrega de cestas básicas, medicamentos e promoção familiar).

## Estrutura

| Pasta | Conteúdo |
|---|---|
| `assets/` | Logo da ONG (usada no papel timbrado) |
| `anamnese-familiar/` | Anamnese Social da Família — cadastro, diagnóstico socioeconômico e Plano de Desenvolvimento Familiar |
| `prestacao-de-contas/` | Prestação de Contas Mensal (relatório + balancete, 2 páginas), Política de Boas Práticas e Declaração de Conformidade |
| `timbrado.py` | Papel timbrado e helpers compartilhados pelos geradores (python-docx) |
| `prestacao-de-contas/` | Prestação de Contas Mensal (balancete), Política de Boas Práticas e Declaração de Conformidade |
| `timbrado.py` | Papel timbrado e helpers compartilhados pelos geradores (python-docx) |

## Anamnese Social da Família

Formulário de 4 páginas, inspirado no Prontuário SUAS / Cadastro Único (CRAS),
voltado ao acompanhamento de famílias em vulnerabilidade e à construção, junto
com o voluntário, de um projeto de geração de renda (trabalho, emprego,
empreendimento ou agropecuário).

Arquivos:

- `Anamnese_Social_Familiar_Santidade_Integral.docx` — versão editável
- `Anamnese_Social_Familiar_Santidade_Integral.pdf` — versão para impressão
- `gerar_anamnese.py` — gerador (python-docx); regenera o DOCX a partir do código

Para regenerar:

```bash
cd anamnese-familiar
python3 gerar_anamnese.py
soffice --headless --convert-to pdf Anamnese_Social_Familiar_Santidade_Integral.docx
```

## Prestação de contas

Três documentos, gerados por `prestacao-de-contas/gerar_prestacao_contas.py`:

| Documento | Uso | Páginas |
|---|---|---|
| `Prestacao_de_Contas_Mensal` | Relatório mensal de execução físico-financeira com balancete e conciliação (impressão mensal) | 2 |
| `Politica_de_Boas_Praticas` | Norma interna de governança, transparência e fidelidade, com fundamento bíblico (honra, dignidade, fidelidade) | 2 |
| `Declaracao_de_Conformidade` | Atestado assinado das boas práticas aplicadas no período — anexo da prestação de contas | 1 |

Referências de nicho: ITG 2002 (R1) do CFC (entidades sem finalidade de lucros), Lei 13.019/2014 (MROSC) e LGPD.

Para regenerar:

```bash
cd prestacao-de-contas
python3 gerar_prestacao_contas.py   # gera DOCX e PDF dos três documentos
```

## Prestação de contas

Conjunto de três documentos, alinhados à ITG 2002 (R1) do Conselho Federal de
Contabilidade (entidades sem fins lucrativos) e ao MROSC (Lei 13.019/2014)
quando houver recurso público, com fundamento bíblico em honra, dignidade e
fidelidade:

- `Prestacao_de_Contas_Mensal` — impresso mensalmente (2 páginas): execução
  física, receitas, despesas, estoque de doações, balancete conciliado,
  verificação de conformidade e pareceres.
- `Politica_de_Boas_Praticas` — documento normativo com os princípios e as
  práticas de governança, gestão financeira, atendimento e transparência.
- `Declaracao_de_Conformidade` — atestado periódico, assinado pela diretoria e
  Conselho Fiscal, das boas práticas aplicadas no período.

Para regenerar:

```bash
cd prestacao-de-contas
python3 gerar_prestacao_contas.py   # gera DOCX e PDF dos três documentos
```
