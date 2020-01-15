---
title: Tempo de execução no arquivo de manifesto
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120632"
---
# <a name="runtime-element"></a>Elemento Runtime

Este recurso está em visualização. Elemento filho do [`<Runtimes>`](runtime.md) elemento. Este elemento facilita o compartilhamento de dados globais e chamadas de função entre as funções personalizadas do Excel e o painel de tarefas do seu suplemento.

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Contido em

-[Tempos](runtimes.md)

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Lifetime = "Long"**  |  Sim  | Deve sempre ser listado como longo se você quiser que as funções personalizadas do Excel funcionem enquanto o painel de tarefas do seu suplemento estiver fechado. |
|  **resid**  |  Sim  | Se usado para funções personalizadas do Excel, `resid` o deve apontar `TaskPaneAndCustomFunction.Url`para. |

## <a name="see-also"></a>Confira também

-[Tempo](runtime.md)
