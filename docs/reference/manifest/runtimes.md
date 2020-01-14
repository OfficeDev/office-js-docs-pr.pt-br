---
title: Tempos de execução no arquivo de manifesto
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111174"
---
# <a name="runtimes-element"></a>Elemento de runtimes

Este recurso está em visualização. Especifica o tempo de execução do suplemento e permite que as funções personalizadas e o painel de tarefas compartilhem dados globais e façam chamadas de função entre si. Deve seguir o `<Host>` elemento no seu arquivo de manifesto.

**Tipo de suplemento:** Painel de tarefas

## <a name="syntax"></a>Sintaxe

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Runtime**     | Sim |  O tempo de execução do suplemento, geralmente usado com funções personalizadas do Excel.

## <a name="see-also"></a>Confira também

-[Tempos](runtimes.md)
