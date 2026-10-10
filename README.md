# Santidade Integral em Cristo — Documentos

Repositório de documentos institucionais da ONG **Santidade Integral em Cristo**
(ação social: entrega de cestas básicas, medicamentos e promoção familiar).

## Estrutura

| Pasta | Conteúdo |
|---|---|
| `assets/` | Logo da ONG (usada no papel timbrado) |
| `anamnese-familiar/` | Anamnese Social da Família — cadastro, diagnóstico socioeconômico e Plano de Desenvolvimento Familiar |

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
