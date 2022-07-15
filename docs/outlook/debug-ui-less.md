---
title: Comandos de função de depuração em suplementos do Outlook
description: Saiba como depurar comandos de função em suplementos do Outlook.
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6189824fd526d48321b355c9b306fa5ef732f411
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797586"
---
# <a name="debug-function-commands-in-outlook-add-ins"></a>Comandos de função de depuração em suplementos do Outlook

> [!NOTE]
> A técnica neste artigo só pode ser usada em um computador de desenvolvimento do Windows. Se você estiver desenvolvendo em um Mac, consulte comandos [de função de depuração](../testing/debug-function-command.md).

Este artigo descreve como usar a Extensão do Depurador de Suplementos do Office no Visual Studio Code para depurar [comandos de função](add-in-commands-for-outlook.md#run-a-function-command). Os comandos de função são iniciados por meio de um botão de comando de suplemento na faixa de opções. Para obter mais informações sobre comandos de suplemento, consulte [comandos de suplemento do Outlook](add-in-commands-for-outlook.md).

Este artigo pressupõe que você já tenha um projeto de suplemento que gostaria de depurar. Para criar um suplemento com um comando de função para praticar a depuração, siga as etapas no Tutorial: Criar um suplemento de [composição de mensagem do Outlook](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Marcar seu suplemento para depuração

Se você usou o gerador [Yeoman para Suplementos do Office](../develop/yeoman-generator-overview.md) para criar seu projeto de suplemento, vá para Configurar e execute a seção [do depurador](#configure-and-run-the-debugger) posteriormente neste artigo. Quando você executa `npm start` para compilar o suplemento e iniciar o servidor local, `UseDirectDebugger` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` o comando também define o valor da chave do Registro para marcar o suplemento para depuração.

Caso contrário, se você usou outra ferramenta para criar seu suplemento, execute as etapas a seguir.

1. Navegue até a `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` chave do Registro. Substitua `[Add-in ID]` pelo **\<Id\>** manifesto do suplemento.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Defina o valor da `UseDirectDebugger` chave como `1`.

## <a name="configure-and-run-the-debugger"></a>Configurar e executar o depurador

Agora que você habilitou a depuração no suplemento, está pronto para configurar e executar o depurador. Para obter instruções sobre como fazer isso, selecione uma das opções a seguir que se aplicam ao controle webview. Para obter informações sobre como determinar qual controle de modo de exibição da Web é usado em seu computador de desenvolvimento, consulte [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).

- Se o suplemento for executado no controle de modo de exibição da Web inserido do Edge Legacy (EdgeHTML), consulte a Extensão do Depurador de Suplementos do [Microsoft Office para Visual Studio Code](../testing/debug-with-vs-extension.md).

- Se o suplemento for executado no controle de modo de exibição da Web inserido do Microsoft Edge Chromium (WebView2), consulte [Depurar suplementos no Windows usando o Visual Studio Code e o Microsoft Edge WebView2 (](../testing/debug-desktop-using-edge-chromium.md)baseados em Chromium).

## <a name="see-also"></a>Confira também

- [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md)
- [Visão geral da depuração de Suplementos do Office](../testing/debug-add-ins-overview.md)
- [Depurar seu suplemento do Outlook baseado em evento](debug-autolaunch.md)
