🧠 Junior Dev - Automação SAP SE16N - EKKO
💡 Visão Geral

Este projeto é uma automação desenvolvida em Python que integra com o SAP GUI Scripting para extrair dados da transação SE16N, tabela EKPO, e exportá-los automaticamente em formato .TXT.

A automação conta com uma interface gráfica moderna, com tons pastéis, onde o usuário pode iniciar o processo de exportação com apenas um clique.

🖥️ Funcionalidades

Conecta automaticamente ao SAP GUI via COM interface (win32com.client);

Executa a transação SE16N e consulta a tabela EKPO;

Exporta os dados da tela de resultados;

Salva o arquivo automaticamente em uma pasta específica do usuário;

Converte o arquivo .XLS exportado em .TXT;

Interface gráfica intuitiva com os botões Iniciar e Voltar.

🎨 Interface

A interface foi desenvolvida em Tkinter, com design minimalista e cores suaves.
No centro da tela, há o título “Junior Dev” e dois botões:

🟩 Iniciar → executa a automação SAP

🔙 Voltar → fecha o aplicativo
