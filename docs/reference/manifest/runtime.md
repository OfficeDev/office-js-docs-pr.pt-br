---
title: Tempo de execução no arquivo de manifesto (versão prévia)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283881"
---
# <a name="runtime-element-preview"></a>Elemento Runtime (visualização)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Elemento filho do [`<Runtimes>`](runtimes.md) elemento. Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Tipo de suplemento:** Painel de tarefas

> [!IMPORTANT]
<<<<<<< o tempo de execução compartilhado HEAD está atualmente em versão prévia e está disponível apenas no Excel no Windows. Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).

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
|  **Lifetime = "Long"**  |  Sim  | Deve ser `long` sempre se você quiser usar um tempo de execução compartilhado para o suplemento do Excel. |
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do suplemento. O `resid` deve corresponder a `id` um atributo de `Url` um elemento no `Resources` elemento. |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
