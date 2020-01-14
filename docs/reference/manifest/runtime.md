---
title: Tempo de execução no arquivo de manifesto
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111167"
---
# <a name="runtime-element"></a>Elemento Runtime

Este recurso está em visualização. Elemento filho do [`<Runtimes>`](runtime.md) elemento. Este elemento facilita o compartilhamento de dados globais e chamadas de função entre as funções personalizadas do Excel e o painel de tarefas do seu suplemento. 

## <a name="contained-in"></a>Contido em

-[Tempos](runtimes.md)

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Lifetime = "Long"**  |  Sim  | Deve sempre ser listado como longo se você quiser que as funções personalizadas do Excel funcionem enquanto o painel de tarefas do seu suplemento estiver fechado. |
|  **resid**  |  Sim  | Se usado para funções personalizadas do Excel, `resid` o deve apontar `TaskPaneAndCustomFunction.Url`para. |

## <a name="see-also"></a>Confira também

-[Tempo](runtime.md)
