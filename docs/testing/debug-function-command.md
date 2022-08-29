---
title: Depurar um comando de função com um runtime não compartilhado
description: Saiba como depurar comandos de função.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: d2be148c05f88837610b8563c2e61618d1c37775
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423199"
---
# <a name="debug-a-function-command-with-a-non-shared-runtime"></a>Depurar um comando de função com um runtime não compartilhado

> [!IMPORTANT]
> Se o suplemento estiver configurado para usar um [runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md) compartilhado, depure o código por trás do comando de função, assim como faria com o código por trás de um painel de tarefas. Consulte [Depurar Suplementos do Office](debug-add-ins-overview.md) e observe que um comando de função em um suplemento com um [runtime](runtimes.md#shared-runtime) compartilhado não é  um caso especial, conforme descrito neste artigo. 

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com comandos [de função](../design/add-in-commands.md#types-of-add-in-commands).

Os comandos de função não têm uma interface do usuário, portanto, um depurador não pode ser anexado ao processo no qual a função é executada na área de trabalho do Office. (Os suplementos do Outlook que estão sendo desenvolvidos no Windows são uma exceção a isso. Consulte [os comandos de função de depuração nos suplementos do Outlook no Windows](#debug-function-commands-in-outlook-add-ins-on-windows) mais adiante neste artigo.) Portanto, os comandos de função, em suplementos com um runtime não compartilhado, devem ser depurados Office na Web em que a função é executada no processo geral do navegador. Use as etapas a seguir.

1. Fazer sideload do suplemento no Office na Web e, em seguida, selecione o botão ou item de menu que executa o comando de função. Isso é necessário para carregar o arquivo de código para o comando de função. 
1. Abra as ferramentas de desenvolvedor do navegador. Isso geralmente é feito pressionando F12. O depurador nas ferramentas é anexado ao processo do navegador.
1. Aplique pontos de interrupção ao código conforme necessário para o comando de função.
1. Execute novamente o comando de função. O processo é interrompido em seus pontos de interrupção. 

> [!TIP]
> Para obter informações mais detalhadas, consulte [Suplementos de depuração no Office na Web](debug-add-ins-in-office-online.md).

## <a name="debug-function-commands-in-outlook-add-ins-on-windows"></a>Comandos de função de depuração em suplementos do Outlook no Windows

Se o computador de desenvolvimento for o Windows, há uma maneira de depurar um comando de função na área de trabalho do Outlook. Consulte [comandos de função de depuração em suplementos do Outlook](../outlook/debug-ui-less.md).

## <a name="see-also"></a>Confira também

- [Runtimes em Suplementos do Office](runtimes.md)
