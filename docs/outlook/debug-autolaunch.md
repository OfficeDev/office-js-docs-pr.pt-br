---
title: Depurar seu suplemento do Outlook baseado em eventos
description: Saiba como depurar o suplemento do Outlook que implementa a ativação baseada em eventos.
ms.topic: article
ms.date: 10/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: e8065c454bbe1587a6e5b7189a4522c229e9aed1
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664669"
---
# <a name="debug-your-event-based-outlook-add-in"></a>Depurar seu suplemento do Outlook baseado em eventos

Este artigo fornece diretrizes de depuração à medida que você implementa [a ativação baseada em eventos](autolaunch.md) em seu suplemento. O recurso de ativação baseado em evento foi introduzido no [conjunto de requisitos 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), com eventos adicionais agora disponíveis em conjuntos de requisitos subsequentes. Para obter mais informações, consulte [Eventos com suporte](autolaunch.md#supported-events).

> [!IMPORTANT]
> Essa funcionalidade de depuração só tem suporte no Outlook no Windows com uma assinatura do Microsoft 365.

Este artigo discute as principais etapas para habilitar a depuração.

- [Marcar o suplemento para depuração](#mark-your-add-in-for-debugging)
- [Configurar Visual Studio Code](#configure-visual-studio-code)
- [Anexar Visual Studio Code](#attach-visual-studio-code)
- [Depurar](#debug)

Se você usou o Gerador yeoman para suplementos do Office para criar seu projeto de suplemento (por exemplo, fazendo o [passo a passo de ativação baseado em evento](autolaunch.md)), siga a opção **Gerador Criado com Yeoman** ao longo deste artigo. Caso contrário, siga as **Outras etapas** . Visual Studio Code deve ser pelo menos a versão 1.56.1.

## <a name="mark-your-add-in-for-debugging"></a>Marcar seu suplemento para depuração

1. Defina a chave `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`do registro . Substitua `[Add-in ID]` pela ID do suplemento do manifesto.

    - **Manifesto XML**: use o valor do **\<Id\>** elemento filho do elemento raiz **\<OfficeApp\>** .
    - **Manifesto do Teams (versão prévia)**: use o valor da propriedade "id" do objeto anônimo `{ ... }` raiz.

    **Criado com o gerador Yeoman**: em uma janela de linha de comando, navegue até a raiz da pasta de suplemento e execute o seguinte comando.

    ```command&nbsp;line
    npm start
    ```

    Além de criar o código e iniciar o servidor local, esse comando deve definir a chave do `UseDirectDebugger` registro para esse suplemento como `1`.

    **Outro**: Adicionar a chave do `UseDirectDebugger` registro em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Substitua `[Add-in ID]` pelo **\<Id\>** do manifesto de suplemento. Defina a chave do registro como `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Inicie o Outlook ou reinicie-o se ele já estiver aberto.
1. Compor uma nova mensagem ou compromisso. A caixa de diálogo Manipulador baseado em Eventos de Depuração deve ser exibida. *Não* interaja com a caixa de diálogo ainda.

    ![A caixa de diálogo Manipulador baseado em Eventos de Depuração no Windows.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Configurar Visual Studio Code

### <a name="created-with-yeoman-generator"></a>Criado com o gerador Yeoman

1. De volta à janela de linha de comando, abra Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. Em Visual Studio Code, abra o arquivo **./.vscode/launch.json** e adicione o seguinte trecho à sua lista de configurações. Salve suas alterações.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a>Outros

1. Crie uma nova pasta chamada **Depuração** (talvez na pasta **Área de Trabalho** ).
1. Abra o Visual Studio Code.
1. Vá para **Pasta Abrir Arquivo** > , navegue até a pasta que você acabou de criar e escolha **Selecionar Pasta**.
1. Na Barra de Atividades, selecione **Executar e Depurar** (Ctrl+Shift+D).

    ![O ícone Executar e Depurar na Barra de Atividades.](../images/vs-code-debug.png)

1. Selecione o link **criar um arquivo launch.json** .

    ![O link localizado na opção Executar e Depurar para criar um arquivo launch.json no Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. Na lista suspensa **Selecionar Ambiente** , **selecione Borda: Iniciar** para criar um arquivo launch.json.
1. Adicione o trecho a seguir à sua lista de configurações. Salve suas alterações.

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a>Anexar Visual Studio Code

1. Para localizar o **bundle.js** do suplemento, abra a pasta a seguir no Windows Explorer e pesquise os suplementos **\<Id\>** (encontrados no manifesto).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Abra a pasta prefixada com essa ID e copie seu caminho completo. Em Visual Studio Code, abra **bundle.js** dessa pasta. O padrão do caminho do arquivo deve ser o seguinte:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Coloque pontos de interrupção no bundle.js em que você deseja que o depurador pare.
1. Na lista suspensa **DEBUG** , selecione **Depuração Direta** e, em seguida, selecione o ícone **Iniciar Depuração** .

    ![A opção Depuração Direta selecionada nas opções de configuração na lista suspensa Visual Studio Code Depuração.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Depurar

1. Depois de confirmar que o depurador está anexado, retorne ao Outlook e, na caixa **de diálogo Manipulador baseado em Eventos de Depuração** , escolha **OK** .

1. Agora você pode atingir seus pontos de interrupção no Visual Studio Code, permitindo que você depure seu código de ativação baseado em eventos.

## <a name="stop-debugging"></a>Parar a depuração

Para parar a depuração para o restante da sessão de área de trabalho atual do Outlook, na caixa **de diálogo Manipulador baseado em Eventos de Depuração** , escolha **Cancelar**. Para habilitar novamente a depuração, reinicie a área de trabalho do Outlook.

Para impedir que a caixa de diálogo **manipulador baseada em eventos de depuração** apareça e pare de depurar para sessões subsequentes do Outlook, exclua a chave do registro associada ou defina seu valor como `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.

## <a name="see-also"></a>Confira também

- [Configurar o suplemento do Outlook para ativação baseada em eventos](autolaunch.md)
- [Depurar seu suplemento com o log do tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows)
