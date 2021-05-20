---
title: Depurar seu complemento Outlook baseado em eventos (pré-visualização)
description: Aprenda a depurar seu Outlook complemento que implementa ativação baseada em eventos.
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555250"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a>Depurar seu complemento Outlook baseado em eventos (pré-visualização)

Este artigo fornece orientação de depuração à medida que você implementa [a ativação baseada](autolaunch.md) em eventos em seu complemento. O recurso de ativação baseado em eventos está atualmente em pré-visualização.

> [!IMPORTANT]
> Esse recurso de depuração só é suportado para visualização em Outlook em Windows com uma assinatura Microsoft 365. Para obter mais informações, consulte a [depuração do Preview para a](#preview-debugging-for-the-event-based-activation-feature) seção de recursos de ativação baseada em eventos neste artigo.

Neste artigo, discutimos as etapas-chave para permitir a depuração.

- [Marque o complemento para depuração](#mark-your-add-in-for-debugging)
- [Configure Visual Studio Code](#configure-visual-studio-code)
- [Anexar Visual Studio Code](#attach-visual-studio-code)
- [depurar](#debug)

Você tem várias opções para criar seu projeto de complemento. Dependendo da opção que você está usando, as etapas podem variar. Onde este é o caso, se você usou o gerador Yeoman para Office Add-ins para criar seu projeto de complementação (por exemplo, fazendo o [passo a passo de ativação baseado](autolaunch.md)em eventos ), então siga as etapas do **escritório,** siga as **outras** etapas. Visual Studio Code deve ser pelo menos a versão 1.56.1.

## <a name="preview-debugging-for-the-event-based-activation-feature"></a>Depuração de visualização para o recurso de ativação baseado em eventos

Convidamos você a experimentar o recurso de depuração para o recurso de ativação baseado em eventos! Deixe-nos saber seus cenários e como podemos melhorar nos dando feedback através de GitHub (veja a seção **Feedback** no final desta página).

Para visualizar essa capacidade para Outlook em Windows, a compilação mínima necessária é de 16.0.13729.20000. Para ter acesso a Office compilações beta, participe do [programa Office Insider](https://insider.office.com).

## <a name="mark-your-add-in-for-debugging"></a>Marque seu complemento para depuração

1. Defina a chave de registro `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` . `[Add-in ID]` é o **ID** no manifesto add-in.

    **yo office**: Em uma janela de linha de comando, navegue até a raiz da pasta de complementação e execute o seguinte comando.

    ```command&nbsp;line
    npm start
    ```

    Além de construir o código e iniciar o servidor local, este comando deve definir a `UseDirectDebugger` chave de registro para este complemento. `1`

    **Outros:** Adicione a `UseDirectDebugger` chave de registro em `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` . Substitua `[Add-in ID]` pelo **ID** do manifesto de complemento. Defina a chave de registro para `1` .

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Inicie Outlook desktop (ou reinicie Outlook se já estiver aberto).
1. Componha uma nova mensagem ou nomeação. Você deve ver o seguinte diálogo. *Ainda não* interaja com o diálogo.

    ![Captura de tela da caixa de diálogo do manipulador baseado em eventos Debug](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Configure Visual Studio Code

### <a name="yo-office"></a>yo escritório

1. De volta à janela da linha de comando, abra Visual Studio Code.

    ```command&nbsp;line
    code .
    ```

1. Em Visual Studio Code, abra o arquivo **./.vscode/launch.js** e adicione o trecho a seguir à sua lista de configurações. Salve suas alterações.

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

1. Crie uma nova pasta chamada **Depuração** (talvez na pasta **Desktop).**
1. Abra Visual Studio Code.
1. Vá para  >  **File Open Folder**, navegue até a pasta que você acabou de criar e escolha Selecionar **pasta**.
1. Na Barra de Atividades, selecione o item **Depuração** (Ctrl+Shift+D).

    ![Captura de tela do ícone Debug na Barra de Atividades](../images/vs-code-debug.png)

1. Selecione **a criação de um launch.jsno** link de arquivo.

    ![Captura de tela do link para criar um launch.jsno arquivo em Visual Studio Code](../images/vs-code-create-launch.json.png)

1. Na lista suspensa **do Ambiente Select,** selecione **Borda: Inicie** para criar uma launch.jsno arquivo.
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

1. Para encontrar o **bundle.js** do complemento, abra a seguinte pasta no Windows Explorer e pesquise o ID do seu **complemento** (encontrado no manifesto).

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    Abra a pasta prefixada com este ID e copie seu caminho completo. Em Visual Studio Code, abra **bundle.js** dessa pasta. O padrão do caminho do arquivo deve ser o seguinte:

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. Coloque pontos de interrupção em bundle.js onde você quer que o depurador pare.
1. Na **lista suspensa do DEBUG,** selecione o nome **Depuração Direta**, e selecione **Executar**.

    ![Captura de tela de seleção de depuração direta das opções de configuração no Visual Studio Code Debug Dropdown](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>depurar

1. Depois de confirmar que o depurador está conectado, retorne ao Outlook e na caixa de diálogo manipulador baseado em **Evento Debug,** escolha **OK** .

1. Agora você pode acertar seus pontos de interrupção em Visual Studio Code, permitindo que você depure seu código de ativação baseado em eventos.

## <a name="stop-debugging"></a>Pare de depurar

Para parar de depurar o resto da sessão de desktop Outlook atual, na caixa de diálogo do manipulador baseado em **eventos Debug,** escolha **Cancelar**. Para reo enable depurar, reinicie Outlook desktop.

Para evitar que a caixa de diálogo do **manipulador baseado em Eventos de depuração** apareça e pare de depurar sessões de Outlook subsequentes, exclua a tecla de registro associada ou defina seu valor `0` para: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .

## <a name="see-also"></a>Confira também

- [Configure seu Outlook complemento para ativação baseada em eventos](autolaunch.md)
- [Depurar seu suplemento com o log do tempo de execução](../testing/runtime-logging.md#runtime-logging-on-windows)
