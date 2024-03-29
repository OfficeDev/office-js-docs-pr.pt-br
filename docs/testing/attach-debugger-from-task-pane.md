---
title: Anexar um depurador do painel de tarefas
description: Saiba como anexar um depurador no painel de tarefas.
ms.date: 01/27/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0363b7966ab3da11167cb4b0cd324df28fd9efb3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744753"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>Anexar um depurador do painel de tarefas

Em alguns ambientes, um depurador pode ser anexado em um Office que já está em execução. Isso pode ser útil quando você deseja depurar um complemento que já está em preparação ou produção. Se você ainda estiver desenvolvendo e testando o complemento, consulte [Overview of debugging Office Add-ins](debug-add-ins-overview.md).

A técnica descrita neste artigo só pode ser usada quando as seguintes condições são atendidas.

- O complemento está sendo executado em Office no Windows.
- O computador está usando uma combinação de versões Windows e Office que usam o controle webview Edge (Chromium baseado em Chromium), WebView2. Para determinar qual navegador você está usando, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Para iniciar o depurador, escolha o canto superior direito do painel de tarefas para ativar **o menu Personalidade** (conforme mostrado no círculo vermelho na imagem a seguir).

![Captura de tela do menu Anexar Depurador.](../images/attach-debugger.png)

Selecione **Anexar Depurador**. Isso inicia as ferramentas de desenvolvedor Microsoft Edge (Chromium baseadas em Chromium). Use as ferramentas conforme descrito em [Depurar os complementos usando ferramentas de desenvolvedor no Microsoft Edge (Chromium baseados em Chromium)](debug-add-ins-using-devtools-edge-chromium.md).

## <a name="see-also"></a>Confira também

- [Visão geral da depuração de Suplementos do Office](debug-add-ins-overview.md)
