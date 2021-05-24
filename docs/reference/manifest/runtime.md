---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd09abe31ff57eac629c6c61c873c5c886f73f9c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590908"
---
# <a name="runtime-element"></a>Elemento Runtime

Configura seu complemento para usar um tempo de execução javaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução. Filho do [`<Runtimes>`](runtimes.md) elemento.

**Tipo de complemento:** Painel de tarefas, Email

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
| [Override](override.md) | Não | **Outlook**: especifica o local da URL do arquivo JavaScript que Outlook Desktop requer para manipuladores de ponto de extensão [LaunchEvent.](../../reference/manifest/extensionpoint.md#launchevent) **Importante:** no momento, você só pode definir um `<Override>` elemento e ele deve ser do tipo `javascript` .|

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do seu complemento. O `resid` pode ter não mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento. |
|  **lifetime**  |  Não  | O valor padrão `lifetime` para é e não precisa ser `short` especificado. Outlook os complementos usam apenas o `short` valor. Se você quiser usar um tempo de execução compartilhado em um Excel de Excel, de definir explicitamente o valor como `long` . |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
