---
title: Elemento Event no arquivo de manifesto
description: Define um manipulador de eventos em um suplemento.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 3d8e94c10bed214dd976b3048e11328f10f99325
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611544"
---
# <a name="event-element"></a>Elemento Event

Define um manipulador de eventos em um suplemento.

> [!NOTE]
> Para obter informações sobre o suporte e uso, consulte [recurso ao enviar para suplementos do Outlook](../../outlook/outlook-on-send-addins.md).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Tipo](#type-attribute)  |  Sim  | Especifica o evento a ser manipulado. |
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
