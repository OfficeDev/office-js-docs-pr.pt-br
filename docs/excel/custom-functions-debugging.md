---
title: Depuração de funções personalizadas sem interface do usuário
description: Saiba como depurar suas funções personalizadas do Excel que não usam um painel de tarefas.
ms.date: 05/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1268aa07d160970fda12f8fccbe88e0427b246fc
ms.sourcegitcommit: 81f6018ac9731ff73e36d30f5ff10df21504c093
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2022
ms.locfileid: "65891925"
---
# <a name="ui-less-custom-functions-debugging"></a>Depuração de funções personalizadas sem interface do usuário

Este artigo aborda a depuração  apenas para funções personalizadas que não usam um painel de tarefas ou outros elementos de interface do usuário (funções personalizadas sem interface do usuário).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

No Windows:

- [Depurador da área de trabalho do Excel e do VS Code (Visual Studio Code)](#use-the-vs-code-debugger-for-excel-desktop)
- [Depurador do Excel na Web e do VS Code](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Ferramentas do Excel na Web e do navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

No Mac:

- [Ferramentas do Excel na Web e do navegador](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [Linha de comando](#use-the-command-line-tools-to-debug)

> [!NOTE]
> Para simplificar, este artigo mostra a depuração no contexto de uso do Visual Studio Code para editar, executar tarefas e, em alguns casos, usar a exibição de depuração. Se você estiver usando uma ferramenta de linha de comando ou editor diferente, consulte as instruções de linha [de comando no](#commands-for-building-and-running-your-add-in) final deste artigo.

## <a name="requirements"></a>Requisitos

Esse processo de depuração **funciona** apenas para funções personalizadas sem interface do usuário, que não usam um painel de tarefas ou outros elementos da interface do usuário. Para criar uma função personalizada sem interface do usuário, siga as etapas no tutorial Criar funções [personalizadas no Excel](../tutorials/excel-tutorial-create-custom-functions.md) e remova todos os elementos do painel de tarefas e da interface do usuário instalados pelo gerador [Yeoman para suplementos do Office](../develop/yeoman-generator-overview.md).

Observe que esse processo de depuração não é compatível com projetos de funções personalizadas usando um [runtime compartilhado](../develop/configure-your-add-in-to-use-a-shared-runtime.md).

Esse processo de depuração não funciona com projetos criados com o projeto de suplemento do **Office** que contém a opção somente de manifesto no gerador Yeoman. Os scripts referenciados posteriormente neste artigo não são instalados com essa opção. Para depurar um suplemento criado com essa opção, consulte as instruções em um desses artigos, conforme apropriado. 

- [Depurar suplementos usando ferramentas de desenvolvedor no Microsoft Edge (baseado em Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md)
- [Depurar suplementos usando ferramentas de desenvolvedor no Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md)
- [Depurar Suplementos do Office em um Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Usar o depurador do VS Code para Excel Desktop

Você pode usar o VS Code para depurar funções personalizadas sem interface do usuário no Office Excel na área de trabalho.

> [!NOTE]
> A depuração da área de trabalho para Mac não está disponível, mas pode ser obtida usando as ferramentas do navegador e a linha de comando para [depurar o Excel na Web](#use-the-command-line-tools-to-debug).

### <a name="run-your-add-in-from-vs-code"></a>Executar o suplemento do VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [VS Code](https://code.visualstudio.com/).
1. Escolha **Executar Tarefa > Terminal e** digite ou selecione **Inspecionar**. Isso monitorará e recriará as alterações de arquivo.
1. Escolha **Terminal > Executar Tarefa e** digite ou selecione **Servidor de Desenvolvimento**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

1. Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o modo de exibição de depuração.
1. No menu suspenso Executar, escolha **Área de Trabalho do Excel (Funções Personalizadas)**.
1. Selecione **F5** (ou **Executar -> Iniciar Depuração** no menu) para iniciar a depuração. Uma nova pasta de trabalho do Excel será aberta com seu suplemento já com sideload e pronto para uso.

### <a name="start-debugging"></a>Iniciar a depuração

1. No VS Code, abra o arquivo de script de código-fonte (**functions.js** **ou functions.ts**).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que use sua função personalizada.

Neste ponto, a execução será interrompida na linha de código em que você define o ponto de interrupção. Agora você pode percorrer seu código, definir relógios e usar todos os recursos de depuração do VS Code necessários.

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Usar o depurador do VS Code para Excel no Microsoft Edge

Você pode usar o VS Code para depurar funções personalizadas sem interface do usuário no Excel no navegador Microsoft Edge. Para usar o VS Code com o Microsoft Edge, você deve instalar a extensão [Microsoft Edge DevTools para Visual Studio Code](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension).

### <a name="run-your-add-in-from-vs-code"></a>Executar o suplemento do VS Code

1. Abra a pasta do projeto raiz de funções personalizadas no [VS Code](https://code.visualstudio.com/).
2. Escolha **Executar Tarefa > Terminal e** digite ou selecione **Inspecionar**. Isso monitorará e recriará as alterações de arquivo.
3. Escolha **Terminal > Executar Tarefa e** digite ou selecione **Servidor de Desenvolvimento**.

### <a name="start-the-vs-code-debugger"></a>Iniciar o depurador do VS Code

1. Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o modo de exibição de depuração.
1. Nas opções de Depuração, escolha **Office Online (Edge Chromium)**.
1. Abra o Excel no navegador Microsoft Edge e crie uma nova pasta de trabalho.
1. Escolha **Compartilhar** na faixa de opções e copie o link para a URL desta nova pasta de trabalho.
1. Selecione **F5** (ou **selecione Executar > Iniciar Depuração** no menu) para iniciar a depuração. Um prompt será exibido, que solicita a URL do documento.
1. Cole a URL da pasta de trabalho e pressione Enter.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Selecione a **guia** Inserir na faixa de opções e, na seção **Suplementos** , escolha **Suplementos do Office**.
2. Na caixa **de diálogo Suplementos do Office** , selecione a guia **MEUS SUPLEMENTOS** , escolha Gerenciar Meus **Suplementos** e Carregue **Meu Suplemento**.
  
    ![A caixa de diálogo Suplementos do Office com uma lista suspensa no canto superior direito lendo "Gerenciar meus suplementos" e uma lista suspensa abaixo dela com a opção "Carregar Meu Suplemento".](../images/office-add-ins-my-account.png)

3. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.
  
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>Definir pontos de interrupção

1. No VS Code, abra o arquivo de script de código-fonte (**functions.js** **ou functions.ts**).
2. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.
3. Na pasta de trabalho do Excel, insira uma fórmula que use sua função personalizada.

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>Usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas no Excel na Web

Você pode usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas sem interface do usuário no Excel na Web. As etapas a seguir funcionam para Windows e macOS.

### <a name="run-your-add-in-from-visual-studio-code"></a>Executar o suplemento no Visual Studio Code

1. Abra a pasta do projeto raiz de funções personalizadas [no Visual Studio Code (VS Code)](https://code.visualstudio.com/).
2. Escolha **Executar Tarefa > Terminal e** digite ou selecione **Inspecionar**. Isso monitorará e recriará as alterações de arquivo.
3. Escolha **Terminal > Executar Tarefa e** digite ou selecione **Servidor de Desenvolvimento**.

### <a name="sideload-your-add-in"></a>Realizar o sideload do seu suplemento

1. Abra [o Office na Web](https://office.live.com/).
2. Abra uma nova pasta de trabalho do Excel.
3. Abra a **guia** Inserir na faixa de opções e, na seção **Suplementos** , escolha **Suplementos do Office**.
4. Na caixa **de diálogo Suplementos do Office** , selecione a guia **MEUS SUPLEMENTOS** , escolha Gerenciar Meus **Suplementos** e Carregue **Meu Suplemento**.
  
    ![A caixa de diálogo Suplementos do Office com uma lista suspensa no canto superior direito lendo "Gerenciar meus suplementos" e uma lista suspensa abaixo dela com a opção "Carregar Meu Suplemento".](../images/office-add-ins-my-account.png)

5. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.
  
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> Depois de fazer o sideload para o documento, ele permanecerá com sideload sempre que você abrir o documento.

### <a name="start-debugging"></a>Iniciar a depuração

1. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
2. Nas ferramentas de desenvolvedor, abra o arquivo de script de código-fonte usando **Cmd+P** ou **Ctrl+P** (**functions.js** **ou functions.ts**).
3. [Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada. 

Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

## <a name="use-the-command-line-tools-to-debug"></a>Usar as ferramentas de linha de comando para depurar

Se você não estiver usando o VS Code, poderá usar a linha de comando (como Bash ou PowerShell) para executar o suplemento. Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel na Web. Você não pode depurar a versão da área de trabalho do Excel usando a linha de comando.

1. Na linha de comando, execute para `npm run watch` observar e recompilar quando ocorrerem alterações de código.
2. Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução do relógio.)

3. Se você quiser iniciar o suplemento na versão da área de trabalho do Excel, execute o comando a seguir.
  
    `npm run start:desktop`
  
    Ou se você preferir iniciar o suplemento no Excel na Web, execute o comando a seguir.
  
    `npm run start:web -- --document {url}` (onde está `{url}` a URL de um arquivo do Excel no OneDrive ou no SharePoint)
  
    Se o suplemento não realizar o sideload no documento, siga as etapas em [Sideload](#sideload-your-add-in) do suplemento para realizar o sideload do suplemento. Em seguida, prossiga para a próxima seção para iniciar a depuração.
  
4. Abra as ferramentas de desenvolvedor no navegador. Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.
5. Nas ferramentas de desenvolvedor, abra o arquivo de script de código-fonte (**functions.js** **ou functions.ts**). O código de funções personalizadas pode estar localizado próximo ao final do arquivo.
6. No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.

Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações. Atualize o navegador para ver as alterações carregadas.

### <a name="commands-for-building-and-running-your-add-in"></a>Comandos para criar e executar seu suplemento

Há várias tarefas de build disponíveis.

- `npm run watch`: compila para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo
- `npm run build-dev`: compilações para desenvolvimento uma vez
- `npm run build`: builds para produção
- `npm run dev-server`: executa o servidor Web usado para desenvolvimento

Você pode usar as tarefas a seguir para iniciar a depuração na área de trabalho ou online.

- `npm run start:desktop`: inicia o Excel na área de trabalho e o sideload do suplemento.
- `npm run start:web -- --document {url}` (onde `{url}` está a URL de um arquivo do Excel no OneDrive ou no SharePoint): inicia o Excel na Web e faz o sideload do suplemento.
- `npm run stop`: interrompe o Excel e a depuração.

## <a name="next-steps"></a>Próximas etapas

Saiba mais sobre [as práticas de autenticação para funções personalizadas sem interface do usuário](custom-functions-authentication.md).

## <a name="see-also"></a>Confira também

* [Solução de problemas de funções personalizadas](custom-functions-troubleshooting.md)
* [Tratamento de erros para funções personalizadas no Excel](custom-functions-errors.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
