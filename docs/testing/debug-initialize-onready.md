---
title: Depurar os métodos initialize e onReady
description: Saiba como depurar os métodos Office.initialize e Office.onReady.
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed6e69a52f3f4534db075daf62c171d4806e89d4
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797699"
---
# <a name="debug-the-initialize-and-onready-methods"></a>Depurar os métodos initialize e onReady

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com [a inicialização do suplemento do Office](../develop/initialize-add-in.md).

O paradoxal da depuração dos métodos [Office.initialize](/javascript/api/office#office-office-initialize-function(1)) e [Office.onReady](/javascript/api/office#office-office-onready-function(1)) é que um depurador só pode anexar a um processo em execução, mas esses métodos são executados imediatamente à medida que o processo de runtime do suplemento é iniciado, antes que um depurador possa anexar. Na maioria das situações, reiniciar o suplemento depois que um depurador é anexado não ajuda porque reiniciar o suplemento fecha o processo de runtime original e o *depurador* anexado e inicia um novo processo que não tem nenhum depurador anexado.

Felizmente, há uma exceção. Você pode depurar esses métodos usando Office na Web, com as etapas a seguir.

1. Fazer sideload e executar o suplemento no Office na Web. Isso geralmente é feito abrindo o painel de tarefas de um suplemento ou executando um comando [de função](../design/add-in-commands.md#types-of-add-in-commands). *O suplemento é executado no processo geral do navegador, não em um processo separado como faria no Office da área de trabalho.*
1. Abra as ferramentas de desenvolvedor do navegador. Isso geralmente é feito pressionando F12. O depurador nas ferramentas é anexado ao processo do navegador.
1. Aplique pontos de interrupção conforme necessário ao código nos `Office.initialize` métodos `Office.onReady` ou no código.
1. *Reiniciar o painel de tarefas* do suplemento ou o comando de função, assim como você fez na etapa 1. Essa ação não *fecha* o processo do navegador ou o depurador. O `Office.initialize` método `Office.onReady` ou é executado novamente e o processamento para em seus pontos de interrupção.

> [!TIP]
> Para obter informações mais detalhadas, consulte [Suplementos de depuração no Office na Web](debug-add-ins-in-office-online.md). 
