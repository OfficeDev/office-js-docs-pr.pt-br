# <a name="extend-excel-functionality"></a>Estender a funcionalidade do Excel

Além de interagir com o conteúdo em uma pasta de trabalho, um suplemento do Excel pode adicionar comandos de menu ou botões personalizados da faixa de opções, inserir painéis de tarefas, abrir caixas de diálogo e, até mesmo, inserir conteúdo rico baseado na Web diretamente em uma planilha.

## <a name="add-in-commands"></a>Comandos de suplemento

Os comandos de suplemento são elementos da interface de usuário que estendem a interface de usuário do Excel e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão da faixa de opções ou um item a um menu de contexto no Excel. Ao selecionar um comando de suplemento, os usuários iniciam ações como executar código JavaScript ou exibir uma página do suplemento em um painel de tarefas. 

**Comandos de suplemento**

![Comandos de suplemento no Excel](../../images/Excel_add-in_commands_Script-Lab.png)

Para saber mais sobre recursos de comando, plataformas suportadas e práticas recomendadas para o desenvolvimento de comandos de suplemento, confira [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md).

## <a name="task-panes"></a>Painéis de tarefas

Os painéis de tarefas são superfícies de interface que normalmente são exibidas no lado direito da janela no Excel. Os painéis de tarefas concedem aos usuários acesso a controles de interface que executam códigos para modificar o documento do Excel ou exibir dados de uma fonte de dados. 

**Painel de tarefas**

![Suplemento do painel de tarefas no Excel](../../images/Excel_add-in_task_pane_Insights.png)

Para saber mais sobre os painéis de tarefas, confira [Painéis de tarefas nos Suplementos do Office](../design/task-pane-add-ins.md). Para ver uma amostra que implementa um painel de tarefas no Excel, confira [Suplemento do Excel JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).

## <a name="dialog-boxes"></a>Caixas de diálogo

As caixas de diálogo são superfícies que flutuam acima da janela do aplicativo do Excel ativo. Você pode usar caixas de diálogo para tarefas como exibir páginas de entrada que não podem ser abertas diretamente em um painel de tarefas, solicitar que o usuário confirme uma ação ou hospedar vídeos que possam ser muito pequenos se confinados a um painel de tarefas. Para abrir caixas de diálogo no suplemento do Excel, use a [API da Caixa de Diálogo](../../reference/shared/officeui.md).

**Caixa de diálogo**

![Caixa de diálogo do suplemento no Excel](../../images/Excel_add-in_dialog_choose-number.png)

Para saber mais sobre caixas de diálogo e a API da Caixa de Diálogo, confira [Caixas de diálogo nos Suplementos do Office](../design/dialog-boxes.md) e [Usar a API da Caixa de Diálogo em Suplementos do Office](../develop/dialog-api-in-office-add-ins.md).

## <a name="content-add-ins"></a>Suplementos de conteúdo

Os suplementos de conteúdo são superfícies que podem ser inseridas diretamente em documentos do Excel. É possível usar suplementos de conteúdo para inserir objetos sofisticados baseados na Web, como gráficos, visualizações de dados ou mídia em uma planilha ou para conceder aos usuários acesso aos controles de interface que executam código para modificar o documento do Excel ou exibir dados de uma fonte de dados. Use suplementos de conteúdo quando quiser inserir a funcionalidade diretamente no documento.

**Suplemento de conteúdo**

![Suplemento de conteúdo no Excel](../../images/Excel_add-in_content_map.png)

Para saber mais sobre suplementos conteúdos, confira [Suplementos do Office de conteúdo](../design/content-add-ins.md). Para ver um exemplo que implementa um suplemento de conteúdo no Excel, confira [Suplemento de conteúdo do Excel Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) no GitHub.

## <a name="additional-resources"></a>Recursos adicionais

- [Comandos de suplemento para Excel, Word e PowerPoint](../design/add-in-commands.md)
- [Definir comandos de suplemento no manifesto](../develop/define-add-in-commands.md)
- [Exemplos de comandos do suplemento do Office no Github](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/)
- [Painéis de tarefas nos suplementos do Office](../design/task-pane-add-ins.md)
- [Suplemento do Excel: JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Caixas de diálogo em Suplementos do Office](../design/dialog-boxes.md)
- [Usar a API da Caixa de Diálogo em suplementos do Office](../develop/dialog-api-in-office-add-ins.md)
- [Exemplo da API da Caixa de Diálogo do suplemento do Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Suplementos do Office de conteúdo](../design/content-add-ins.md)
- [Suplemento de conteúdo do Excel: Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
