---
title: Tempos de execução no arquivo de manifesto
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554003"
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

- [Runtime](runtime.md)
