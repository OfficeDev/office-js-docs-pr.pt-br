---
ms.date: 04/09/2021
description: Saiba como depurar suas funções personalizadas do Excel que não usam um painel de tarefas.
title: Depuração de funções personalizadas sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: 5b27ca44dbb891c2e1f4ae86175595dc902b74ba
ms.sourcegitcommit: 094caf086c2696e78fbdfdc6030cb0c89d32b585
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/16/2021
ms.locfileid: "51862334"
---
# <a name="ui-less-custom-functions-debugging"></a>Depuração de funções personalizadas sem interface do usuário

Este artigo discute a depuração *apenas* para funções personalizadas que não usam um painel de tarefas ou outros elementos de interface do usuário (funções personalizadas sem interface do usuário). 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

No Windows:
- [Depurador de código Visual Studio e área de trabalho do Excel (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Depurador do Excel na Web e vs Code](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel na Web e ferramentas do navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

No Mac:
- [Excel na Web e ferramentas do navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Para simplificar, este artigo mostra a depuração no contexto de uso do código Visual Studio para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração. Se você estiver usando uma ferramenta de linha de comando ou editor diferente, consulte [as](#commands-for-building-and-running-your-add-in) instruções de linha de comando no final deste artigo.

## <a name="requirements"></a>Requisitos

Esse processo de depuração funciona **apenas** para funções personalizadas sem interface do usuário, que não usam um painel de tarefas ou outros elementos da interface do usuário. Uma função personalizada sem interface do usuário pode ser criada seguindo as etapas no tutorial Criar funções [personalizadas](../tutorials/excel-tutorial-create-custom-functions.md) no Excel e removendo todos os elementos do painel de tarefas e da interface do usuário instalados pelo gerador [Yeoman](https://www.npmjs.com/package/generator-office)para Os Complementos do Office.

Observe que esse processo de depuração não é compatível com projetos de funções personalizadas usando um [tempo de execução compartilhado.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Usar o depurador de código VS para Área de Trabalho do Excel

Você pode usar o VS Code para depurar funções personalizadas sem interface do usuário no Office Excel na área de trabalho.

> [!NOTE]
> A depuração de área de trabalho para o Mac não está disponível, mas pode ser atingida usando as ferramentas do navegador e a linha de comando para [depurar o Excel na Web](#use-the-command-line-tools-to-debug)).

### <a name="run-your-add-in-from-vs-code"></a>Executar o seu complemento a partir do código VS

1. Abra sua pasta de projeto raiz de funções personalizadas em [VS Code](https://code.visualstudio.com/).
2. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**. Isso monitorará e reconstruirá todas as alterações de arquivo.
3. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador de código VS

4. Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o exibição de depuração.
5. No menu suspenso Executar, escolha **Área de Trabalho do Excel (Edge Chromium)**.
6. Selecione **F5** (ou selecione **Executar -> Iniciar Depuração** no menu) para começar a depuração. Uma nova planilha do Excel será aberta com seu complemento já sideload e pronto para uso.

### <a name="start-debugging"></a>Iniciar a depuração

1. Em Vs Code, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**).
2. [Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na planilha excel, insira uma fórmula que usa sua função personalizada.

Neste ponto, a execução será parada na linha de código onde você definirá o ponto de interrupção. Agora você pode passar pelo código, definir relógios e usar todos os recursos de depuração do VS Code necessários.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Usar o depurador de código VS para Excel no Microsoft Edge

Você pode usar o VS Code para depurar funções personalizadas sem interface do usuário no Excel no navegador do Microsoft Edge. Para usar o VS Code com o Microsoft Edge, você deve instalar o [Depurador para a extensão do Microsoft Edge.](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)

### <a name="run-your-add-in-from-vs-code"></a>Executar o seu complemento a partir do código VS

1. Abra sua pasta de projeto raiz de funções personalizadas em [VS Code](https://code.visualstudio.com/).
2. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**. Isso monitorará e reconstruirá todas as alterações de arquivo.
3. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador de código VS

4. Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o exibição de depuração.
5. Nas opções Depurar, escolha **Office Online (Edge Chromium)**.
6. Abra o Excel no navegador do Microsoft Edge e crie uma nova planilha.
7. Escolha **Compartilhar** na faixa de opções e copie o link para a URL dessa nova workbook.
8. Selecione **F5** (ou **selecione Executar > Iniciar Depuração** no menu) para começar a depuração. Um prompt será exibido, que solicita a URL do documento.
9. Colar na URL da pasta de trabalho e pressione Enter.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Selecione a **guia** Inserir na faixa de opções e, na seção **Complementos,** escolha **Office Add-ins**.
2. Na caixa **de diálogo Complementos** do Office, selecione a guia **MEUS ADD-INS,** escolha Gerenciar Meus **Complementos** e, em seguida, **Carregue Meu Add-in**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

3. **Navegue** até o arquivo de manifesto do complemento e selecione **Carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>Definir pontos de interrupção
1. Em Vs Code, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**).
2. [Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na planilha excel, insira uma fórmula que usa sua função personalizada.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Use as ferramentas de desenvolvedor do navegador para depurar funções personalizadas no Excel na Web

Você pode usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas sem interface do usuário no Excel na Web. As etapas a seguir funcionam para Windows e macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Execute o seu add-in do Visual Studio Code

1. Abra sua pasta de projeto raiz de funções personalizadas [no Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**. Isso monitorará e reconstruirá todas as alterações de arquivo.
3. Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Abra [o Office na Web](https://office.live.com/).
2. Abra uma nova planilha do Excel.
3. Abra a **guia** Inserir na faixa de opções e, na seção **Complementos,** escolha **Complementos do Office**.
4. Na caixa **de diálogo Complementos** do Office, selecione a guia **MEUS ADD-INS,** escolha Gerenciar Meus **Complementos** e, em seguida, **Carregue Meu Add-in**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> Depois de fazer sideload no documento, ele permanecerá sideload sempre que você abrir o documento.

### <a name="start-debugging"></a>Iniciar a depuração

1. Abra ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores F12 abrirá as ferramentas de desenvolvedor.
2. Em ferramentas de desenvolvedor, abra seu arquivo de script de código-fonte usando **Cmd+P** ou **Ctrl+P** (**functions.js** **ou functions.ts**).
3. [Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada. 

Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

## <a name="use-the-command-line-tools-to-debug"></a>Usar as ferramentas de linha de comando para depurar

Se você não estiver usando o VS Code, poderá usar a linha de comando (como bash ou PowerShell) para executar o seu complemento. Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel na Web. Não é possível depurar a versão da área de trabalho do Excel usando a linha de comando.

1. Na linha de comando, `npm run watch` execute para observar e reconstruir quando ocorrerem alterações de código.
2. Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução do relógio).

3. Se você quiser iniciar seu complemento na versão da área de trabalho do Excel, execute o seguinte comando
    
    `npm run start:desktop`
    
    Ou se você preferir iniciar seu complemento no Excel na Web execute o seguinte comando
    
    `npm run start:web`
    
    Para o Excel na Web, você também precisa fazer sideload do seu complemento. Siga as etapas em [Sideload your add-in](#sideload-your-add-in) to sideload your add-in. Em seguida, continue até a próxima seção para iniciar a depuração.
    
4. Abra ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores F12 abrirá as ferramentas de desenvolvedor.
5. Em ferramentas de desenvolvedor, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**). Seu código de funções personalizadas pode estar localizado perto do final do arquivo.
6. No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.

Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

### <a name="commands-for-building-and-running-your-add-in"></a>Comandos para criar e executar o seu complemento

Há várias tarefas de com build disponíveis:
- `npm run watch`: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo
- `npm run build-dev`: builds para desenvolvimento uma vez
- `npm run build`: builds para produção
- `npm run dev-server`: executa o servidor Web usado para desenvolvimento

Você pode usar as seguintes tarefas para iniciar a depuração na área de trabalho ou online.
- `npm run start:desktop`: Inicia o Excel na área de trabalho e faz o sideload do seu complemento.
- `npm run start:web`: Inicia o Excel na Web e faz o sideload do seu complemento.
- `npm run stop`: Interrompe o Excel e a depuração.

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre as práticas de autenticação para funções [personalizadas sem interface do usuário.](custom-functions-authentication.md)

## <a name="see-also"></a>Confira também

* [Solução de problemas de funções personalizadas](custom-functions-troubleshooting.md)
* [Tratamento de erros para funções personalizadas no Excel](custom-functions-errors.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
