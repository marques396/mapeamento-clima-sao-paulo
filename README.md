Mapeamento Clima São Paulo
Este projeto automatiza a coleta de dados climáticos para a cidade de São Paulo e outras regiões do Brasil. Originalmente desenvolvido utilizando técnicas de Web Scraping com Selenium, o sistema foi atualizado para consumir dados diretamente da OpenWeatherMap API, tornando a coleta mais rápida, estável e profissional.

⚙️ Funcionalidades
Busca em Tempo Real: Obtém temperatura, umidade e condição climática via API REST.

Interface Gráfica (GUI): Desenvolvida em Python com a biblioteca Tkinter para facilitar a interação.

Persistência de Dados: Salva automaticamente cada consulta em um arquivo Excel (clima.xlsx) para gerar um histórico de medições.

Automação DevOps: O projeto já conta com arquivos de configuração de ambiente como .gitignore e .gitattributes para garantir a integridade do código no controle de versão.

🛠️ Tecnologias Utilizadas
Linguagem: Python 3.14.

Bibliotecas:

requests: Para o consumo da API.

openpyxl: Para manipulação de planilhas Excel.

tkinter: Para a interface visual.

🚀 Como Rodar o Projeto
Instale as dependências necessárias:

PowerShell
pip install requests openpyxl
Obtenha uma chave gratuita em OpenWeatherMap.

Insira sua chave na variável API_KEY dentro do arquivo MAPEAMENTO_CLIMA_SÃO_PAULO.PY.

Execute o script:

PowerShell
python ./DadosClima/MAPEAMENTO_CLIMA_SÃO_PAULO.PY
