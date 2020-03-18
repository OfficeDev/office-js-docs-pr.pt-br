---
ms.date: 07/10/2019
description: Depurar suas funções personalizadas no Excel.
title: Depuração de funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 4abd5f3da58c35485004b17f92b334b133cabd27
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719305"
---
# <a name="custom-functions-debugging"></a>Depuração de funções personalizadas

A depuração de funções personalizadas pode ser realizada por vários meios, dependendo de qual plataforma você está usando.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

No Windows:
- [Depurador de área de trabalho do Excel e Visual Studio (VS Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Web Excel e depurador de código VS](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel na Web e ferramentas de navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

No Mac:
- [Excel na Web e ferramentas de navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Para simplificar, este artigo mostra a depuração no contexto de uso do Visual Studio Code para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração. Se você estiver usando um editor ou uma ferramenta de linha de comando diferente, consulte as [instruções de linha de comando](#commands-for-building-and-running-your-add-in) no final deste artigo.

## <a name="requirements"></a>Requirements

Antes de começar a depurar, você deve usar o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) para criar um projeto de funções personalizadas. Para obter orientação sobre como criar um projeto de funções personalizadas, consulte o [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Usar o depurador de código VS para a área de trabalho do Excel

Você pode usar o VS Code para depurar funções personalizadas no Office Excel na área de trabalho.

> [!NOTE]
> A depuração de área de trabalho do Mac não está disponível, mas pode ser obtida [usando as ferramentas de navegador e a linha de comando para depurar o Excel na Web](#use-the-command-line-tools-to-debug)).

### <a name="run-your-add-in-from-vs-code"></a>Executar seu suplemento de VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).
2. Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

4. Escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.
5. Nas opções de depuração, escolha **área de trabalho do Excel**.
6. Selecione **F5** (ou escolha **debug-> iniciar a depuração** no menu) para iniciar a depuração. Uma nova pasta de trabalho do Excel será aberta com seu suplemento já suplementos foi feito e pronto para uso.

### <a name="start-debugging"></a>Iniciar Depuração

1. No VS Code, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.

Nesse ponto, a execução será interrompida na linha de código em que você definir o ponto de interrupção. Agora você pode percorrer seu código, definir inspeções e usar quaisquer recursos de depuração de código VS necessários.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Usar o depurador de código VS para Excel no Microsoft Edge

Você pode usar o VS Code para depurar funções personalizadas no Excel no navegador Microsoft Edge. Para usar o VS Code com o Microsoft Edge, você deve instalar o depurador para a extensão do [Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .

### <a name="run-your-add-in-from-vs-code"></a>Executar seu suplemento de VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).
2. Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

4. Escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.
5. Nas opções de depuração, escolha **Office Online (Microsoft Edge)**.
6. Abra o Excel no navegador Microsoft Edge e crie uma nova pasta de trabalho.
7. Escolha **compartilhar** na faixa de opções e copie o link para a URL dessa nova pasta de trabalho.
8. Selecione **F5** (ou escolha **debug > iniciar a depuração** no menu) para iniciar a depuração. Um prompt será exibido, solicitando a URL do seu documento.
9. Cole na URL da sua pasta de trabalho e pressione Enter.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Selecione a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.
2. Na caixa de diálogo **suplementos do Office** , selecione a guia **meus suplementos** , escolha **gerenciar meus suplementos**e, em seguida, **carregar meu suplemento**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

3. **Navegue** até o arquivo de manifesto do suplemento e selecione **carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>Definir pontos de interrupção
1. No VS Code, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas no Excel na Web

Você pode usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas no Excel na Web. As etapas a seguir funcionam para o Windows e o macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Executar seu suplemento do Visual Studio Code

1. Abra a pasta do projeto raiz de suas funções personalizadas no [Visual Studio Code (vs Code)](https://code.visualstudio.com/).
2. Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**. Isso irá monitorar e recriar qualquer alteração de arquivo.
3. Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Abra o [Microsoft Office na Web](https://office.live.com/).
2. Abra uma nova pasta de trabalho do Excel.
3. Abra a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.
4. Na caixa de diálogo **suplementos do Office** , selecione a guia **meus suplementos** , escolha **gerenciar meus suplementos**e, em seguida, **carregar meu suplemento**.
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> Depois que você tiver suplementos foi feito para o documento, ele permanecerá suplementos foi feito cada vez que você abrir o documento.

### <a name="start-debugging"></a>Iniciar Depuração

1. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
2. Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte usando **cmd + p** ou **Ctrl + p** (**funções. js** ou **funções. TS**).
3. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada. 

Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

## <a name="use-the-command-line-tools-to-debug"></a>Usar as ferramentas de linha de comando para depurar

Se você não estiver usando o VS, poderá usar a linha de comando (como bash ou PowerShell) para executar o suplemento. Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel na Web. Não é possível depurar a versão da área de trabalho do Excel usando a linha de comando.

1. A partir da linha de `npm run watch` comando, execute para observar e recriar quando ocorrerem alterações de código.
2. Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução da inspeção).

3. Se você deseja iniciar o suplemento na versão da área de trabalho do Excel, execute o seguinte comando
    
    `npm run start:desktop`
    
    Ou se preferir iniciar seu suplemento no Excel na Web, execute o seguinte comando
    
    `npm run start:web`
    
    Para o Excel na Web, você também precisa Sideload seu suplemento. Siga as etapas em [Sideload seu suplemento](#sideload-your-add-in) para Sideload o suplemento. Em seguida, prossiga para a próxima seção para iniciar a depuração.
    
4. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
5. Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**). O código de suas funções personalizadas pode estar localizado próximo ao final do arquivo.
6. No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.

Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

### <a name="commands-for-building-and-running-your-add-in"></a>Comandos para compilar e executar o suplemento

Há várias tarefas de compilação disponíveis:
- `npm run watch`: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo
- `npm run build-dev`: cria para desenvolvimento uma vez
- `npm run build`: compilações para produção
- `npm run dev-server`: executa o servidor Web usado para desenvolvimento

Você pode usar as seguintes tarefas para iniciar a depuração no desktop ou online.
- `npm run start:desktop`: Inicia o Excel na área de trabalho e sideloads seu suplemento.
- `npm run start:web`: Inicia o Excel na Web e sideloads seu suplemento.
- `npm run stop`: Interrompe o Excel e a depuração.

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre as [práticas de autenticação em funções personalizadas](custom-functions-authentication.md). Ou, revise a [arquitetura exclusiva da função personalizada](custom-functions-architecture.md).

## <a name="see-also"></a>Também confira

* [Solução de problemas de funções personalizadas](custom-functions-troubleshooting.md)
* [Tratamento de erros para funções personalizadas no Excel](custom-functions-errors.md)
* [Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário](make-custom-functions-compatible-with-xll-udf.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
