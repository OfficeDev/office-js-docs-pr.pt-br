---
title: Tempo de execução no arquivo de manifesto
description: O elemento Runtime configura seu complemento para usar um tempo de execução JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fa95608d7eff57d68b96ef5b04ec9d33ee63f173
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652241"
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

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do seu complemento. O `resid` pode ter não mais de 32 caracteres e deve corresponder a um atributo de um elemento no `id` `Url` `Resources` elemento. |
|  **lifetime**  |  Não  | O valor padrão `lifetime` para é e não precisa ser `short` especificado. Os complementos do Outlook usam apenas o `short` valor. Se você quiser usar um tempo de execução compartilhado em um complemento do Excel, de definir explicitamente o valor como `long` . |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
- [Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [Configurar seu complemento do Outlook para ativação baseada em eventos](../../outlook/autolaunch.md)
