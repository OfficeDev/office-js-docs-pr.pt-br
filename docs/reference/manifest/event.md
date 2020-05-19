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
# <a name="event-element"></a><span data-ttu-id="a69dc-103">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="a69dc-103">Event element</span></span>

<span data-ttu-id="a69dc-104">Define um manipulador de eventos em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="a69dc-104">Defines an event handler in an add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a69dc-105">Para obter informações sobre o suporte e uso, consulte [recurso ao enviar para suplementos do Outlook](../../outlook/outlook-on-send-addins.md).</span><span class="sxs-lookup"><span data-stu-id="a69dc-105">For information about support and usage, see [On-send feature for Outlook add-ins](../../outlook/outlook-on-send-addins.md).</span></span>

## <a name="attributes"></a><span data-ttu-id="a69dc-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="a69dc-106">Attributes</span></span>

|  <span data-ttu-id="a69dc-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="a69dc-107">Attribute</span></span>  |  <span data-ttu-id="a69dc-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="a69dc-108">Required</span></span>  |  <span data-ttu-id="a69dc-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="a69dc-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a69dc-110">Type</span><span class="sxs-lookup"><span data-stu-id="a69dc-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="a69dc-111">Sim</span><span class="sxs-lookup"><span data-stu-id="a69dc-111">Yes</span></span>  | <span data-ttu-id="a69dc-112">Especifica o evento a ser manipulado.</span><span class="sxs-lookup"><span data-stu-id="a69dc-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="a69dc-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="a69dc-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="a69dc-114">Sim</span><span class="sxs-lookup"><span data-stu-id="a69dc-114">Yes</span></span>  | <span data-ttu-id="a69dc-p101">Especifica o estilo de execução para o manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="a69dc-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="a69dc-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="a69dc-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="a69dc-118">Sim</span><span class="sxs-lookup"><span data-stu-id="a69dc-118">Yes</span></span>  | <span data-ttu-id="a69dc-119">Especifica o nome da função para o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="a69dc-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="a69dc-120">Atributo de tipo</span><span class="sxs-lookup"><span data-stu-id="a69dc-120">Type attribute</span></span>

<span data-ttu-id="a69dc-p102">Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="a69dc-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="a69dc-124">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="a69dc-124">Event type</span></span>  |  <span data-ttu-id="a69dc-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="a69dc-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="a69dc-126">O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.</span><span class="sxs-lookup"><span data-stu-id="a69dc-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="a69dc-127">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="a69dc-127">FunctionExecution attribute</span></span>

<span data-ttu-id="a69dc-128">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="a69dc-128">Required.</span></span> <span data-ttu-id="a69dc-129">DEVE ser definido como `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="a69dc-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="a69dc-130">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="a69dc-130">FunctionName attribute</span></span>

<span data-ttu-id="a69dc-p104">Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="a69dc-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" />
```
