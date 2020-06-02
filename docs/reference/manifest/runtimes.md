---
title: Tempos de execução no arquivo de manifesto
description: O elemento de Runtime especifica o tempo de execução do seu suplemento.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a8598a8f926e6d6905c147f5c554f1d40a692ad9
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471321"
---
# <a name="runtimes-element"></a>Elemento de runtimes

Especifica o tempo de execução do seu suplemento. Filho do [`<Host>`](host.md) elemento. Se o `Runtimes` elemento estiver presente no manifesto, o suplemento usará o navegador Internet Explorer 11 por padrão.

No Excel, esse elemento permite que a faixa de opções, o painel de tarefas e as funções personalizadas usem o mesmo tempo de execução. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

No Outlook, esse elemento habilita a ativação de suplementos baseada em eventos. Para obter mais informações, consulte [Configure Your Outlook Add-in for Event-based Activation](../../outlook/autolaunch.md).

**Tipo de suplemento:** Painel de tarefas, email

> [!IMPORTANT]
> **Excel**: o tempo de execução compartilhado atualmente só está disponível no Excel no Windows.
>
> **Outlook**: o recurso de ativação baseado em eventos está atualmente [em versão prévia](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) e disponível apenas no Outlook na Web. Para obter mais informações, consulte [como visualizar o recurso de ativação baseado em eventos](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
| [Tempo de execução](runtime.md) | Sim |  O tempo de execução do suplemento. |

## <a name="see-also"></a>Confira também

- [Tempo de execução](runtime.md)
