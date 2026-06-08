# filtro_ncm

Filtro automático de NCM com NVE e Ex-tarifário, executado 100% no navegador, sem backend.

## Visão geral

Esta ferramenta permite cruzar informações de NCM com:

- Ex-tarifários vigentes (planilha EX)
- NVE (Nomenclatura de Valor Especificado, planilha NVE)

A aplicação roda inteiramente no browser, lê os arquivos Excel localmente (via [SheetJS](https://docs.sheetjs.com/)[web:29]) e não envia nenhum dado para servidor. Ideal para consultas rápidas, testes de cenários e geração de relatórios em PDF a partir das bases oficiais.

## Funcionalidades

- Upload separado das bases:
  - Planilha de Ex-tarifário (EX)
  - Planilha de NVE
- Funciona com apenas uma ou com as duas bases carregadas
- Busca por NCM com ou sem pontos (por exemplo `8517.62.15` ou `85176215`)
- Pesquisa por **prefixo**:
  - A partir de 4 dígitos já traz resultados
  - Com 8 dígitos, mostra o NCM completo formatado
- Limpeza de texto para remover artefatos comuns de planilha (`x000D`, quebras estranhas etc.)
- Para EX:
  - Cálculo de vigência a partir da data de fim
  - Destaque visual de linhas vencidas
  - Contador de EX vencidos no resultado
- Geração de **relatório em nova aba**, pronto para impressão ou salvar como PDF (via `window.print`)
- Tema claro/escuro com alternância manual

## Como usar

1. Abra a aplicação (GitHub Pages) no navegador.
2. Na tela inicial, clique em **“Começar”**.
3. Na etapa **1. Carregar planilhas**:
   - Em **Base EX**, selecione a planilha de Ex-tarifário (`.xlsx` ou `.xls`);
   - Em **Base NVE**, selecione a planilha de NVE (`.xlsx` ou `.xls`);
   - É possível usar apenas uma das bases, se quiser.
4. Após o carregamento bem-sucedido de pelo menos uma base, clique em **“Iniciar consulta”**.
5. Na etapa **2. Consultar NCM**:
   - Digite ou cole o NCM com ou sem pontos;
   - A busca começa com 4 dígitos, e o NCM é normalizado para apenas números.
6. Veja os resultados nas tabelas de **NVE** e **Ex-tarifário**:
   - Cabeçalhos fixos;
   - Quebra de texto para descrições longas;
   - EX vencidos destacados.
7. Clique em **“Gerar PDF”** para abrir uma nova aba com o relatório:
   - Selecione “Imprimir” / “Salvar como PDF” no navegador.

## Formato esperado das planilhas

A aplicação espera planilhas com **colunas mínimas** (nomes equivalentes são aceitos, com normalização de acentos e espaços):

### Planilha EX

- `NCM`
- `EX`
- `FIM DA VIGÊNCIA`
- `DESCRIÇÃO` (recomendado)

### Planilha NVE

- `NCM`
- `ATRIBUTO`
- `DESCRIÇÃO ATRIBUTO`
- `ESPECIFICACAO`
- `DESCRIÇÃO ESPECIFICAÇÃO`

Os cabeçalhos são normalizados (maiúsculas, sem acentos, sem espaços extras), então variações simples costumam ser toleradas.

## Stack técnica

- **HTML + CSS + JavaScript** em uma única página (`index.html`)
- **SheetJS (XLSX)** para leitura das planilhas Excel no navegador[web:29]
- Nenhum backend, banco de dados ou armazenamento local:
  - Os arquivos são processados em memória;
  - Ao recarregar a página, tudo é descartado.

## Execução local

Para testar localmente:

1. Clone este repositório:

   ```bash
   git clone https://github.com/SEU_USUARIO/SEU_REPO.git
   cd SEU_REPO
   ```

2. Abra o arquivo `index.html` em um servidor estático, por exemplo com a extensão **Live Server** do VS Code[web:25], ou usando Python:

   ```bash
   python -m http.server 8000
   ```

3. Acesse em `http://localhost:8000/index.html`.

> Observação: abrir o arquivo direto como `file://` pode causar bloqueios em recursos ou comportamentos diferentes entre navegadores; prefira um servidor local simples.

## Limitações conhecidas

- A leitura depende de a estrutura das planilhas respeitar as colunas mínimas listadas acima; se algo crítico não for encontrado, aparece uma mensagem de erro.
- A performance foi pensada para planilhas “reais”, mas ainda em ambiente de estação de trabalho (não é um motor de busca massivo).
- A geração de PDF usa a função nativa de impressão do navegador (`window.print`); o resultado final pode variar levemente entre navegadores e sistemas.
