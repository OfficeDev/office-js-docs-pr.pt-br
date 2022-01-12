---
title: Elemento Event no arquivo de manifesto
description: Define um manipulador de eventos em um suplemento.
ms.date: 01/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: fac920fc91abd908d3d159877c0c414bd7fae244
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765889"
---
# <a name="event-element"></a>Elemento Event

Define um manipulador de eventos em um suplemento.

> [!NOTE]
> Para obter informações sobre suporte e uso, consulte Recurso Ao enviar [para Outlook de complementos](../../outlook/outlook-on-send-addins.md).

**Tipo de suplemento:** Email

**Válido somente nestes esquemas VersionOverrides:**

- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

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
