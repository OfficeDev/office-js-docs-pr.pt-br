---
title: Depurar seu suplemento do Outlook baseado em evento
description: Saiba como depurar seu suplemento do Outlook que implementa a ativação baseada em evento.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5d36a23b34132071077e3eb192e562288befb8a5
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797488"
---
# <a name="debug-your-event-based-outlook-add-in"></a>Depurar seu suplemento do Outlook baseado em evento

Este artigo fornece diretrizes de depuração à medida que [você implementa a](autolaunch.md) ativação baseada em eventos em seu suplemento. O recurso de ativação baseada em evento foi introduzido no conjunto de requisitos [1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) com eventos adicionais agora disponíveis na versão prévia. Para obter mais informações, consulte [eventos com suporte](autolaunch.md#supported-events).

> [!IMPORTANT]
> Esse recurso de depuração só tem suporte no Outlook no Windows com uma assinatura do Microsoft 365.

Neste artigo, discutiremos os principais estágios para habilitar a depuração.

- [Marcar o suplemento para depuração](#mark-your-add-in-for-debugging)
- [Configurar Visual Studio Code](#configure-visual-studio-code)
- [Anexar Visual Studio Code](#attach-visual-studio-code)
- [Depurar](#debug)

Se você usou o Gerador Yeoman para Suplementos do Office para criar seu projeto de suplemento (por exemplo, fazendo o passo a passo de ativação baseada em [evento), siga](autolaunch.md) a opção de gerador Criado com **Yeoman** em todo este artigo. Caso contrário, siga **as outras** etapas. Visual Studio Code deve ser pelo menos a versão 1.56.1.

## <a name="mark-your-add-in-for-debugging"></a>Marcar seu suplemento para depuração

1. Defina a chave do Registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`. `[Add-in ID]` é o **\<Id\>** manifesto do suplemento.

    **Criado com o gerador Yeoman**: em uma janela de linha de comando, navegue até a raiz da pasta do suplemento e execute o comando a seguir.

    ```command&nbsp;line
    npm start
    ```

    Além de criar o código e iniciar o servidor local, `UseDirectDebugger` esse comando deve definir a chave do Registro para esse suplemento como `1`.

    **Outros**: adicione a `UseDirectDebugger` chave do Registro em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`. Substitua `[Add-in ID]` pelo **\<Id\>** manifesto do suplemento. Defina a chave do Registro como `1`.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Inicie a área de trabalho do Outlook (ou reinicie o Outlook se ele já estiver aberto).
1. Redigir uma nova mensagem ou compromisso. Você deverá ver a caixa de diálogo a seguir. Não *interaja* com a caixa de diálogo ainda.

    ![Captura de tela da caixa de diálogo do manipulador baseado em Evento de Depuração.](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Configurar Visual Studio Code

### <a name="created-with-yeoman-generator"></a>Criado com o gerador Yeoman

1. De volta à janela de linha de comando, abra Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. No Visual Studio Code, abra o arquivo **./.vscode/launch.json** e adicione o trecho a seguir à sua lista de configurações. Salve suas alterações.

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

### <a name="other"></a>Other

1. Crie uma nova pasta chamada **Depuração** (talvez na pasta **Área de** Trabalho).
1. Abra o Visual Studio Code.
1. Vá para **Abrir** > **Pasta de Arquivo**, navegue até a pasta que você acabou de criar e escolha **Selecionar Pasta**.
1. Na Barra de Atividades, selecione o item **Depurar** (Ctrl+Shift+D).

    ![Captura de tela do ícone Depurar na Barra de Atividades.](../images/vs-code-debug.png)

1. Selecione o **link criar um arquivo launch.json** .

    ![Captura de tela do link para criar um arquivo launch.json no Visual Studio Code.](../images/vs-code-create-launch.json.png)

1. Na lista **suspensa Selecionar Ambiente** , selecione **Edge: Iniciar** para criar um arquivo launch.json.
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

1. Para localizar o nome dobundle.js, abra **a** pasta a seguir no Windows Explorer **\<Id\>** e pesquise o suplemento (encontrado no manifesto).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Abra a pasta prefixada com essa ID e copie seu caminho completo. Em Visual Studio Code, abra **bundle.js** nessa pasta. O padrão do caminho do arquivo deve ser o seguinte:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Coloque pontos de interrupção bundle.js onde você deseja que o depurador pare.
1. Na lista **suspensa DEBUG** , selecione o nome **Depuração Direta** e, em seguida, **selecione Executar**.

    ![Captura de tela da seleção de Depuração Direta nas opções de configuração na lista suspensa Visual Studio Code Depuração.](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Depurar

1. Depois de confirmar que o depurador está anexado, retorne ao Outlook e, na caixa de diálogo manipulador baseado em evento de **depuração** , escolha **OK** .

1. Agora você pode atingir seus pontos de interrupção Visual Studio Code, permitindo que você depure seu código de ativação baseado em evento.

## <a name="stop-debugging"></a>Parar a depuração

Para interromper a depuração para o restante da sessão atual da área de trabalho do Outlook, na caixa de diálogo manipulador baseado em Evento de **Depuração** , escolha **Cancelar**. Para reabilitar a depuração, reinicie a área de trabalho do Outlook.

Para impedir que a caixa de diálogo do manipulador baseado em evento de **depuração** apareça e interrompa a depuração para sessões subsequentes do Outlook, exclua a chave do Registro associada ou defina seu valor como `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.

## <a name="see-also"></a>Confira também

- [Configurar seu suplemento do Outlook para ativação baseada em evento](autolaunch.md)
- [Depurar seu suplemento com o log do tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows)
