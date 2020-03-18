---
title: Tempo de execução no arquivo de manifesto (versão prévia)
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para sua faixa de opções, painel de tarefas e funções personalizadas.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717926"
---
# <a name="runtime-element-preview"></a>Elemento Runtime (visualização)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Elemento filho do [`<Runtimes>`](runtimes.md) elemento. Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Tipo de suplemento:** Painel de tarefas

> [!IMPORTANT]
> O tempo de execução compartilhado está atualmente em versão prévia e só está disponível no Excel no Windows. Para experimentar os recursos de visualização, você precisará ingressar no [Office Insider](https://insider.office.com/).

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

## <a name="see-also"></a>Também confira

- [Tempos de execução](runtimes.md)
