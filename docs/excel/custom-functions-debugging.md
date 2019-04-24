---
ms.date: 03/13/2019
description: Depurar suas funções personalizadas no Excel.
title: Depuração de funções personalizadas (visualização)
localization_priority: Normal
ms.openlocfilehash: 08563ef630ebc457219c4c622328b84d13e6acab
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448748"
---
# <a name="custom-functions-debugging-preview"></a>Depuração de funções personalizadas (visualização)

A depuração de funções personalizadas pode ser realizada por vários meios, dependendo de qual plataforma você está usando.

No Windows:
- [Depurador de área de trabalho do Excel e Visual Studio (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [O Excel online e o depurador de código VS](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [Excel online e ferramentas de navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [Linha de comando](#use-the-command-line-tools-to-debug)

No Mac:
- [Excel online e ferramentas de navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [Linha de comando](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> Para simplificar, este artigo mostra a depuração no contexto de uso do Visual Studio Code para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração. Se você estiver usando um editor ou uma ferramenta de linha de comando diferente, consulte as [instruções de linha de comando](#use-the-command-line-tools-to-debug) no final deste artigo.

## <a name="requirements"></a>Requirements

Antes de começar a depurar, você deve criar um projeto de suplemento de funções personalizadas usando o gerador de Yo Office e garantiu que você tenha certificados autoassinados confiáveis para o seu projeto. Para obter instruções sobre como criar um projeto, consulte o [tutorial funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md). Para obter instruções sobre como confiar em certificados, consulte [adicionando certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Usar o depurador de código VS para a área de trabalho do Excel

Você pode usar o VS Code para depurar funções personalizadas no Office Excel na área de trabalho.

> [!NOTE]
> A depuração de área de trabalho do Mac não está disponível, mas pode ser obtida [usando as ferramentas de navegador para depurar o Excel online](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online).

### <a name="run-your-add-in-from-vs-code"></a>Executar seu suplemento de VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).
2. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**. 

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

4. Escolha **Exibir _GT_ Debug** ou Enter **Ctrl + Shift + D** para alternar para o modo de depuração.
5. Nas opções de depuração, escolha **área de trabalho do Excel**.
6. Selecione **F5** (ou escolha **debug-> iniciar depuração** no menu) para iniciar a depuração. Uma nova pasta de trabalho do Excel será aberta com seu suplemento já suplementos foi feito e pronto para uso.

### <a name="start-debugging"></a>Iniciar Depuração

1. No VS Code, abra o arquivo de script do código-fonte (funções. js ou funções. TS).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.

Nesse ponto, a execução será interrompida na linha de código em que você definir o ponto de interrupção. Agora você pode percorrer seu código, definir inspeções e usar quaisquer recursos de depuração de código VS necessários.

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a>Usar o depurador de código VS para o Excel online no Microsoft Edge

Você pode usar o VS Code para depurar funções personalizadas no Excel online no navegador Microsoft Edge. Para usar o VS Code com o Microsoft Edge, você deve instalar o depurador para a extensão do [Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .

### <a name="run-your-add-in-from-vs-code"></a>Executar seu suplemento de VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).
2. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**. 

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

4. Escolha **Exibir _GT_ Debug** ou Enter **Ctrl + Shift + D** para alternar para o modo de depuração.
5. Nas opções de depuração, escolha **Office Online (borda)**.
6. Abra o Excel online usando o navegador do Microsoft Edge, abra o Excel online e crie uma nova pasta de trabalho.
7. Escolha **compartilhar** na faixa de opções e copie o link para a URL dessa nova pasta de trabalho.
8. Selecione **F5** (ou escolha **depurar > iniciar depuração** no menu) para iniciar a depuração. Um prompt será exibido, solicitando a URL do seu documento.
9. Cole na URL da sua pasta de trabalho e pressione Enter.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento   

1. Selecione a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.
2. Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

3.  **Navegue** até o arquivo de manifesto do suplemento e selecione **carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>Definir pontos de interrupção
1. No VS Code, abra o arquivo de script do código-fonte (funções. js ou funções. TS).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a>Usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online

Você pode usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online. As etapas a seguir funcionam para o Windows e o macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Executar seu suplemento do Visual Studio Code

1. Abra a pasta do projeto raiz de suas funções personalizadas no [Visual Studio Code (vs Code)](https://code.visualstudio.com/).
2. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**. 

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento   

1. Abra o [Microsoft Office Online](https://office.live.com/).
2. Abra uma nova pasta de trabalho do Excel.
3. Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.
4. Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5.  **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> Depois que você tiver suplementos foi feito para o documento, ele permanecerá suplementos foi feito cada vez que você abrir o documento.

### <a name="start-debugging"></a>Iniciar Depuração

1. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
2. Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte usando **cmd + p** ou **Ctrl + p** (funções. js ou funções. TS).
3. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada. 

Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

## <a name="use-the-command-line-tools-to-debug"></a>Usar as ferramentas de linha de comando para depurar

Se você não estiver usando o VS, poderá usar a linha de comando (como bash ou PowerShell) para executar o suplemento. Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel online. Não é possível depurar a versão da área de trabalho do Excel usando a linha de comando.

1. A partir da linha de `npm run watch` comando, execute para observar e recriar quando ocorrerem alterações de código.
2. Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução da inspeção).

3. Se você deseja iniciar o suplemento na versão da área de trabalho do Excel, execute o seguinte comando
    
    `npm run start desktop`
    
    Ou se preferir iniciar seu suplemento no Excel online, execute o seguinte comando
    
    `npm run start web`
    
    Para o Excel online, você também precisa Sideload seu suplemento. Siga as etapas em [Sideload seu suplemento](#sideload-your-add-in) para Sideload o suplemento. Em seguida, prossiga para a próxima seção para iniciar a depuração.
    
4. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
5. Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte (funções. js ou funções. TS). O código de suas funções personalizadas pode estar localizado próximo ao final do arquivo.
6. No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.

Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

### <a name="commands-for-building-and-running-your-add-in"></a>Comandos para compilar e executar o suplemento

Há várias tarefas de compilação disponíveis:
- `npm run watch`: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo
- `npm run build-dev`: cria para desenvolvimento uma vez
- `npm run build`: compilações para produção
- `npm run dev-server`: executa o servidor Web usado para desenvolvimento

Você pode usar as seguintes tarefas para iniciar a depuração no desktop ou online.
- `npm run start desktop`: Inicia o Excel na área de trabalho e sideloads seu suplemento.
- `npm run start web`: Inicia o Excel online e o sideloads do suplemento.
- `npm run stop`: Interrompe o Excel e a depuração.

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Tempo de execução de funções personalizadas do Excel](custom-functions-runtime.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Log de alteração de funções personalizadas](custom-functions-changelog.md)
* [Tutorial de funções personalizadas do Excel](../tutorials/excel-tutorial-create-custom-functions.md)
