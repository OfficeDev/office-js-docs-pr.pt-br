---
title: Elemento Event no arquivo de manifesto
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: eda895b01e106d67eef70f199be64086e9372bef
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432736"
---
# <a name="event-element"></a>Elemento Event

Define um manipulador de eventos em um suplemento.

> [!NOTE] 
> O elemento `Event` no momento só tem suporte pelo Outlook na Web no Office 365.

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