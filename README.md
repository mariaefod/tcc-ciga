# TCC - Docs as Code na Administração Pública

Projeto de monografia da Pós-graduação em Inovação, Estratégia e Gestão Pública Avançada da Fundação João Pinheiro, escrito em Markdown.

## Estrutura

- `projeto/projeto.md`: texto do projeto
- `slides/projeto.md`: apresentação do projeto
- `assets/`: imagens usadas no projeto e nos slides
- `src/gerar.py`: gera o Word a partir do `projeto.md`
- `src/estilo.py`: formatação do Word, conforme o manual de normalização da FJP
- `src/sumario.py`: preenche o sumário do Word (usado no GitHub Actions)
- `modelos/`: modelos de projeto da FJP

## Instalação

Requer Python 3.13 e Poetry.

```Shell
poetry install
```

## Gerar o Word

```Shell
poetry run python src/gerar.py
```

O arquivo é salvo em `projeto/Projeto - Maria Eduarda.docx`. Localmente, o sumário fica vazio: ele é preenchido no push para a `main`.

## Ver os slides

```Shell
poetry run task serve
```

Depois, abra http://localhost:8000.

## Publicação automática

A cada push na `main`:

- `gerar.yml`: gera o Word com o sumário preenchido e faz o commit do arquivo na `main`
- `slides.yml`: publica os slides em https://mariaefod.github.io/tcc-ciga/

## Como escrever no Markdown

- `#`, `##` e `###`: títulos
- parágrafo que começa com `Tabela`, `Quadro`, `Figura` ou `Gráfico`: título da ilustração
- parágrafo que começa com `Fonte:`: fonte da ilustração
- `X` em uma célula de quadro: célula pintada (cronograma)
- `>` no início da linha: citação longa
- `![](../assets/imagem.png)`: imagem
- `&nbsp;`: parágrafo vazio
