---
title: Elemento Event no arquivo de manifesto
description: Define um manipulador de eventos em um suplemento.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 80f21d1819e3d7e335389070ccac0db583026045
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275704"
---
# <a name="event-element"></a>Elemento Event

Define um manipulador de eventos em um suplemento.

> [!NOTE]
> Para obter informações sobre o suporte e uso, consulte [recurso ao enviar para suplementos do Outlook](../../outlook/outlook-on-send-addins.md).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Type](#type-attribute)  |  Sim  | Especifica o evento a ser manipulado. |
|  [FunctionExecution](#functionexecution-attribute)  |  Sim  | Especifica o estilo de execução para o manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte. |
|  [FunctionName](#functionname-attribute)  |  Sim  | Especifica o nome da função para o manipulador de eventos. |

### <a name="type-attribute"></a>Atributo de tipo

Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.

|  Tipo de evento  |  Descrição  |
|:-----|:-----|
|  `ItemSend`  |  O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.  |

### <a name="functionexecution-attribute"></a>Atributo FunctionExecution

Obrigatório. DEVE ser definido como `synchronous`.

### <a name="functionname-attribute"></a>Atributo FunctionName

Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
