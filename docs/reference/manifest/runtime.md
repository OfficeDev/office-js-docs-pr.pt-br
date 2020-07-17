---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para seus vários componentes, por exemplo, faixa de opções, painel de tarefas, funções personalizadas.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159364"
---
# <a name="runtime-element-preview"></a>Elemento Runtime (visualização)

Configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que vários componentes sejam executados no mesmo tempo de execução. Filho do [`<Runtimes>`](runtimes.md) elemento.

No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos. Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).

**Tipo de suplemento:** Painel de tarefas, email

> [!IMPORTANT]
> **Outlook**: a ativação baseada em evento está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web. Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do suplemento. O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento. |
|  **marca**  |  Não  | O valor padrão para `lifetime` é `short` e não precisa ser especificado. Os suplementos do Outlook usam apenas o `short` valor. Se você quiser usar um tempo de execução compartilhado em um suplemento do Excel, defina explicitamente o valor como `long` . |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
