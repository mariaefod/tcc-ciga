# Projeto de pesquisa: Docs as Code na Administração Pública

Projeto de pesquisa da pós-graduação da Fundação João Pinheiro, escrito em Markdown.

## Arquivos

- `projeto/projeto.md`: texto do projeto; os dados da capa ficam no topo, entre `---`
- `src/estilo.py`: formatação do documento (fonte, tamanhos, espaçamentos, margens e capa)
- `src/gerar.py`: converte o `projeto.md` em `projeto.docx`

## Como gerar o documento

```
poetry install
poetry run python src/gerar.py
```

A cada push na `main`, o GitHub Actions também gera o `projeto.docx`,
disponível na aba Actions, em Artifacts.

## Regras de formatação

- `#`, `##` e `###` viram títulos
- parágrafo que começa com `Tabela ` vira título da tabela
- parágrafo que começa com `Fonte:` vira fonte da tabela, em tamanho 10
- parágrafos depois de `# REFERÊNCIAS` ficam alinhados à esquerda, com espaçamento simples
