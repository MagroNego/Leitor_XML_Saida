# Leitor XML

Aplicação local para transformar XMLs completos de NF-e em relatórios de itens e retenções de PIS/COFINS. A versão 2.0 usa o navegador, com modo claro/escuro, prévia, busca, paginação e exportação.

## Iniciar no Windows

1. Instale Python 3.10 ou superior, com **Add Python to PATH** marcado.
2. Extraia a pasta completa do projeto.
3. Execute **Instalar_Excel_Windows.bat** uma vez para habilitar XLSX. Essa instalação precisa de internet; o uso do leitor não precisa.
4. Execute **Iniciar_Windows.bat**. O navegador abre automaticamente.
5. Mantenha a janela do servidor aberta. Para encerrar, pressione Ctrl+C nela.

Sem openpyxl, o aplicativo funciona e exporta CSV; o botão Excel fica desativado.

## Linux / execução manual

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python Relatorio_saidas.py
```

O endereço padrão é `http://127.0.0.1:8765`. Se a porta estiver ocupada, o aplicativo usa uma porta livre e informa o endereço. Use `--no-browser` para não abrir o navegador e `--port 9000` para escolher outra porta.

## Fluxo

- **Importar XMLs:** selecionar arquivos, selecionar pasta com subpastas ou arrastar arquivos XML. Arraste arquivos, não pastas; para pastas, use o botão correspondente.
- **Relatórios:** conferir itens ou retenções e buscar em todas as colunas.
- **Processamento:** conferir arquivos importados, duplicatas e erros, com exportação do registro.
- **Excel:** abas Itens, Retenções e Processamento; números para cálculos e códigos/chaves como texto.
- **CSV:** separador `;`, UTF-8 com BOM e vírgula decimal. Na importação pelo Excel, defina chaves, códigos, NCM e CST como texto.

Uma nova importação substitui os resultados atuais, após confirmação. Interromper preserva os arquivos já processados. Duplicatas são ignoradas pela chave. Retenções zeradas não são incluídas. A busca atual também filtra a exportação de itens e retenções; a aba Processamento inclui todos os registros.

## Dados e segurança

Os XMLs são lidos pelo servidor no próprio computador, exclusivamente em `127.0.0.1`. Não há envio a serviços externos, telemetria, banco de dados ou gravação automática dos XMLs. Documentos e resultados ficam em memória durante a execução. **Exporte os relatórios antes de encerrar o servidor.** Atualizar a página preserva os dados enquanto o servidor estiver ativo.

A API exige token aleatório de sessão e valida Host/Origin. O XML tem limite de 20 MB e DTD/entidades são rejeitados. A sessão admite até 250 MB de arquivos recebidos; o consumo real de memória pode ser maior devido à expansão do XML e dos relatórios. Textos do XML são exibidos sem HTML e exportados com proteção contra interpretação como fórmulas.

## Escopo

Lê NF-e completa (`NFe` ou `nfeProc`), com ou sem namespace, itens, ICMS, CSOSN, IPI, PIS, COFINS e retenções no grupo `retTrib`. Preserva precisão decimal dos valores unitários no relatório CSV e diferencia tributação por quantidade. A precisão numérica do Excel segue o limite do próprio Excel.

Não consulta SEFAZ, não manifesta notas, não verifica assinatura/autorização/cancelamento e não processa CT-e, NFS-e ou XML de evento/resumo como NF-e completa. Esta versão não exporta NFref, ICMS-ST, FCP, DIFAL ou os campos da reforma tributária. É um relatório de leitura, não uma apuração fiscal.

## Verificação

```bash
python -m unittest discover -s tests -v
```

O pacote inclui código-fonte e iniciadores. Não inclui executável Windows compilado; o `.exe` da versão 1.0 não deve ser usado para abrir a versão 2.0.
