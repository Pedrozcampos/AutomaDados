# AutoDados 
É a ferramenta que criei para fazer procedimentos com arquivos do excel usando uma base e confrontando com demais arquivo. 

# Funcionamento
- Quando você seleciona um arquivo ele pede pra você digitar uma palavra chave (TERMO) onde ele vai procurar e selecionar a linha que contém esse termo e fazer o devido confronto com a base. ele usa SQLite para ter eficiência no processamento e conseguir um resultado melhor.

# AutoDados (GUI)
- Nosso layout para o gui para ficar uma ferramenta mais profissional e fácil de utilizar.

# Procedimentos
- Temos algumas funções na nossa primeira classe onde:
- *tratar_dados*: Limpa os valores monetários e converte para números.
- *validar_col*: Serve para validar se as colunas obrigatorias estão presentes.
- *proc_adicionar*: Procedimento de adicionar o arquivo na ferramenta com sua devida palavra chave.
- *proc_limpar*: Remove todos arquivos da ferramenta e reinicia o banco de dados na mémoria.
- *atualizar_fila*: Atualiza a lista visual conforme é adicionado arquivos.
- *processar_sql*: Faz a busca SQL filtrando utilizando o termo e une os dados para confronto.
- *exportar_exceel*: Gera o arquivo final ( EXCEL ) com todos os dados incluindo o resultado com sua formatação adequada.

# Tecnologias Utilizadas
- *Python*.
- *SQLite3*: É a o motor do banco de dados.
- *Pandas*: Para manipular dados.
- *CustomTkinter*: Para criação do gui.
- *Openpyxl*: Para formatar e estilização do relatório do Excel.
