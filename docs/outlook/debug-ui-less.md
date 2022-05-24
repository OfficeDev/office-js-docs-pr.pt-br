---
title: Depurar seu suplemento sem interface Outlook interface do usuário
description: Saiba como depurar seu suplemento sem interface Outlook interface do usuário.
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 33aa36f86b7a163e650a23296b4c35aca7cb5492
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650705"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>Depurar seu suplemento sem interface Outlook interface do usuário

Este artigo descreve como usar Office Extensão do Depurador de Suplementos no Visual Studio Code para depurar suplementos sem interface do usuário [Outlook usuário](add-in-commands-for-outlook.md#executing-a-javascript-function). As ações de suplemento sem interface do usuário são iniciadas por meio de um botão de comando de suplemento na faixa de opções. Para obter mais informações sobre comandos de suplemento, consulte [comandos de suplemento para Outlook](add-in-commands-for-outlook.md).

Este artigo pressupõe que você já tenha um projeto de suplemento que gostaria de depurar. Para criar um suplemento sem interface do usuário para praticar a depuração, siga as etapas no Tutorial: Criar uma composição de mensagem [Outlook suplemento](../tutorials/outlook-tutorial.md).

## <a name="mark-your-add-in-for-debugging"></a>Marcar seu suplemento para depuração

Se você usou o gerador [Yeoman para Office Suplementos](../develop/yeoman-generator-overview.md) para criar seu projeto de suplemento, vá para Configurar e execute a seção [do depurador](#configure-and-run-the-debugger) mais adiante neste artigo. Quando você executa `npm start` para compilar o suplemento e iniciar o servidor local, `UseDirectDebugger` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` o comando também define o valor da chave do Registro para marcar o suplemento para depuração.

Caso contrário, se você usou outra ferramenta para criar seu suplemento, execute as etapas a seguir.

1. Navegue até a `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` chave do Registro. Substitua `[Add-in ID]` pela **ID** do manifesto do suplemento.

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Defina o valor da `UseDirectDebugger` chave como `1`.

## <a name="configure-and-run-the-debugger"></a>Configurar e executar o depurador

Agora que você habilitou a depuração no suplemento, está pronto para configurar e executar o depurador. Para obter instruções sobre como fazer isso, selecione uma das opções a seguir que se aplicam ao runtime.

- Se o suplemento for executado no runtime do WebView, consulte Microsoft Office Extensão do [Depurador de Suplementos para Visual Studio Code](../testing/debug-with-vs-extension.md).

- Se o suplemento for executado no runtime do Microsoft Edge Chromium WebView2, consulte [os suplementos de depuração no Windows usando Visual Studio Code e Microsoft Edge WebView2 (baseados em Chromium)](../testing/debug-desktop-using-edge-chromium.md).

## <a name="see-also"></a>Confira também

- [Comandos de suplemento para o Outlook](add-in-commands-for-outlook.md)
- [Visão geral da depuração de Suplementos do Office](../testing/debug-add-ins-overview.md)
- [Depurar seu suplemento baseado em Outlook evento](debug-autolaunch.md)
