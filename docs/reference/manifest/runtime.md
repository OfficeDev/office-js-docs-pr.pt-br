---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, fita, painel de tarefas, funções personalizadas.
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: c59e5a23e53940aea46c758d710b4a455cb5c0cc
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555301"
---
# <a name="runtime-element"></a>Elemento runtime

Configura seu complemento para usar um tempo de execução JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução. Filho do [`<Runtimes>`](runtimes.md) elemento.

**Tipo de complemento:** Painel de tarefas, Correio

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

- [Tempos de execução](runtimes.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Substituição](override.md) (visualização) | Não | **Outlook**: Especifica a localização do URL do arquivo JavaScript que Outlook Desktop requer para manipuladores de [pontos de extensão LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent-preview) **Importante**: No momento, você só pode definir um `<Override>` elemento e ele deve ser de tipo `javascript` .|

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid**  |  Sim  | Especifica a localização da URL da página HTML para o seu complemento. O `resid` pode não ter mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento. |
|  **vida**  |  Não  | O valor padrão para `lifetime` é e não precisa ser `short` especificado. Outlook adicionais usam apenas o `short` valor. Se você quiser usar um tempo de execução compartilhado em um complemento Excel, defina explicitamente o valor para `long` . |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configure seu Outlook complemento para ativação baseada em eventos](../../outlook/autolaunch.md)
