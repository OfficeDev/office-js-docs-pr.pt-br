---
title: Tempos de execução no arquivo de manifesto
description: O elemento Runtimes especifica o tempo de execução do seu complemento.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 74bb2b432f46d5876601052003e20ff843e13b06
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104823"
---
# <a name="runtimes-element"></a>Elemento Runtimes

Especifica o tempo de execução do seu complemento. Filho do [`<Host>`](host.md) elemento.

> [!NOTE]
> Ao executar no Office no Windows, seu complemento usa o navegador Internet Explorer 11.

No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução. Para saber mais, confira Configurar seu complemento do Excel para usar um tempo de execução [JavaScript compartilhado.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

No Outlook, esse elemento habilita a ativação de um complemento baseado em eventos. Para saber mais, confira [Configurar seu complemento do Outlook para ativação baseada em eventos.](../../outlook/autolaunch.md)

**Tipo de complemento:** Painel de tarefas, Email

> [!IMPORTANT]
> **Outlook**: o recurso de ativação baseada em eventos está atualmente em [visualização](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e só está disponível no Outlook na Web e no Windows. Para obter mais informações, [consulte Como visualizar o recurso de ativação baseada em eventos.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

[Host](host.md)

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Sim |  O tempo de execução do seu complemento. |

## <a name="see-also"></a>Confira também

- [Runtime](runtime.md)
