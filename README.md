# Excel-Automation
Este script Python automatiza a separação de dados de uma planilha do Excel, organizando-os em abas separadas com base no bairro de cada entrada. Ele lê uma planilha chamada Bairros.xlsx, identifica os bairros presentes na coluna correspondente e transfere os dados para novas abas dentro do arquivo.

📌 Funcionalidades

Verifica se cada bairro já possui uma aba específica.

Cria novas abas para bairros que ainda não existem.

Copia os dados da "Base de Dados" para a aba correspondente ao bairro.

Mantém a formatação dos cabeçalhos e das células originais.

Gera um novo arquivo chamado Bairros2.xlsx com os dados organizados.

🚀 Como Usar

1️⃣ Pré-requisitos

Python 3.8+

Biblioteca openpyxl

Instale a biblioteca necessária com:

pip install openpyxl

2️⃣ Coloque seu arquivo na pasta correta

O script espera um arquivo chamado Bairros.xlsx na mesma pasta onde ele será executado.

3️⃣ Execute o script

python organizar_bairros.py

Após a execução, será gerado um novo arquivo chamado Bairros2.xlsx, contendo os dados organizados por bairro.

🛠️ Estrutura do Código

add_sheet(bairro, archive_bairros, header_style) → Cria uma nova aba para o bairro, se necessário, e adiciona os cabeçalhos.

transfer_data(OG_sheet, dest_sheet, OG_linha) → Copia os dados da planilha original para a aba correspondente.

Loop principal → Percorre todas as linhas da "Base de Dados", verificando e transferindo os dados corretamente.
