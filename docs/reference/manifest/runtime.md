---
title: Tempo de execução no arquivo de manifesto
description: O elemento de tempo de execução configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para sua faixa de opções, painel de tarefas e funções personalizadas.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217757"
---
# <a name="runtime-element"></a>Elemento Runtime

Elemento filho do [`<Runtimes>`](runtimes.md) elemento. Este elemento configura seu suplemento para usar um tempo de execução de JavaScript compartilhado para que a faixa de opções, o painel de tarefas e as funções personalizadas, todos sejam executados no mesmo tempo de execução. Para obter mais informações, consulte [Configurar o suplemento do Excel para usar um tempo de execução do JavaScript compartilhado](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Tipo de suplemento:** Painel de tarefas

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
|  **Lifetime = "Long"**  |  Sim  | Deve ser sempre `long` se você quiser usar um tempo de execução compartilhado para o suplemento do Excel. |
|  **resid**  |  Sim  | Especifica o local da URL da página HTML do suplemento. O `resid` deve corresponder a um `id` atributo de um `Url` elemento no `Resources` elemento. |

## <a name="see-also"></a>Confira também

- [Tempos de execução](runtimes.md)
