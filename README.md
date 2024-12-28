# Readme - Script de Exportação de Dados da Fraga

## Descrição

Este script automatiza a busca de informações detalhadas sobre peças e modelos na API da **Fraga**. Ele utiliza uma lista de modelos de uma planilha para buscar compatibilidades, referências cruzadas e outras informações gerais. Após a busca, o script organiza os dados em **três planilhas Excel**, cada uma contendo informações específicas:

1. **Informações Gerais**: Detalhes técnicos sobre os produtos, como marca, número de peça, dimensões, etc.  
2. **Referências Cruzadas**: Lista de referências cruzadas (substituições ou equivalentes).  
3. **Compatibilidades**: Modelos de veículos compatíveis com os produtos encontrados.

---

## Pré-requisitos

1. **Node.js instalado** no ambiente.
2. **Bibliotecas necessárias**: Execute o comando abaixo para instalar as dependências:
   ```bash
   npm install
   ```
   Principais bibliotecas utilizadas:
   - `node-fetch`: Para realizar requisições HTTP.
   - `exceljs`: Para manipular arquivos Excel.

3. **Planilha de entrada**:
   - Deve estar no formato **Excel** (`.xlsx`).
   - Contendo uma coluna de modelos com o **nome especificado no script** (por padrão, coluna `A` da aba `Planilha1`).
   - Nome do arquivo: `modelos_de_pecas.xlsx`.

4. **Credenciais de login da Fraga**:
   - Insira os dados de **e-mail e senha** no array `logins` dentro do código:
     ```javascript
     let logins = [{ email: "seu_email@dominio.com", password: "sua_senha" }];
     ```

---

## Como o Script Funciona

1. **Leitura da Planilha**:
   - O script lê a coluna especificada da planilha `modelos_de_pecas.xlsx`.

2. **Autenticação na API da Fraga**:
   - Faz login usando as credenciais fornecidas para obter um token de acesso.

3. **Busca de Dados**:
   - Para cada modelo listado na planilha, o script faz consultas na API para coletar:
     - Informações detalhadas do produto.
     - Referências cruzadas.
     - Veículos compatíveis.

4. **Exportação para Excel**:
   - Os resultados são salvos em três arquivos Excel:
     - `Produtos Calpen - Infos Gerais.xlsx`
     - `Produtos Calpen - Referencias Cruzada.xlsx`
     - `Produtos Calpen - Compatibilidades.xlsx`

---

## Como Executar

1. **Configure a planilha de entrada**:
   - Certifique-se de que a planilha `modelos_de_pecas.xlsx` está corretamente preenchida.

2. **Insira as credenciais de login**:
   - Edite o arquivo para adicionar os dados de e-mail e senha no array `logins`.

3. **Execute o script**:
   - Utilize o comando abaixo para rodar o script:
     ```bash
     node nome_do_arquivo.js
     ```

4. **Verifique os resultados**:
   - Os arquivos exportados estarão disponíveis na mesma pasta do script.

---

## Observações

- O script inclui um mecanismo para atualizar automaticamente o token de autenticação a cada 20 minutos.
- Para evitar perda de dados em execuções longas, considere usar um banco de dados como armazenamento intermediário.
- Caso os arrays gerados sejam muito grandes, o script pode consumir muita memória. Ajuste a lógica ou processe os dados em partes, se necessário.
