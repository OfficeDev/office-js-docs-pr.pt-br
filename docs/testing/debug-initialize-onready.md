---
title: Depurar as funções initialize e onReady
description: Saiba como depurar as funções Office.initialize e Office.onReady.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dca551d8a016e7aad16cfdc02590f0a51455852
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423248"
---
# <a name="debug-the-initialize-and-onready-functions"></a>Depurar as funções initialize e onReady

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com [a inicialização do suplemento do Office](../develop/initialize-add-in.md).

O paradoxal da depuração das funções [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) e [Office.onReady](/javascript/api/office#office-office-onready-function(1)) é que um depurador só pode anexar a um processo em execução, mas essas funções são executadas imediatamente à medida que o processo de runtime do suplemento é iniciado, antes que um depurador possa anexar. Na maioria das situações, reiniciar o suplemento depois que um depurador é anexado não ajuda porque reiniciar o suplemento fecha o processo de runtime original e o *depurador* anexado e inicia um novo processo que não tem nenhum depurador anexado.

Felizmente, há uma exceção. Você pode depurar essas funções usando Office na Web, com as etapas a seguir.

1. Fazer sideload e executar o suplemento no Office na Web. Isso geralmente é feito abrindo o painel de tarefas de um suplemento ou executando um comando [de função](../design/add-in-commands.md#types-of-add-in-commands). *O suplemento é executado no processo geral do navegador, não em um processo separado como faria no Office da área de trabalho.*
1. Abra as ferramentas de desenvolvedor do navegador. Isso geralmente é feito pressionando F12. O depurador nas ferramentas é anexado ao processo do navegador.
1. Aplique pontos de interrupção conforme necessário ao código na `Office.initialize` função `Office.onReady` ou no código.
1. *Reiniciar o painel de tarefas* do suplemento ou o comando de função, assim como você fez na etapa 1. Essa ação não *fecha* o processo do navegador ou o depurador. A `Office.initialize` função ou `Office.onReady` é executada novamente e o processamento para em seus pontos de interrupção.

> [!TIP]
> Para obter informações mais detalhadas, consulte [Suplementos de depuração no Office na Web](debug-add-ins-in-office-online.md).

## <a name="see-also"></a>Confira também

- [Runtimes em Suplementos do Office](runtimes.md)
