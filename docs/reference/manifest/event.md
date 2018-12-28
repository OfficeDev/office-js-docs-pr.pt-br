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
# <a name="event-element"></a><span data-ttu-id="cd19c-102">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="cd19c-102">Event element</span></span>

<span data-ttu-id="cd19c-103">Define um manipulador de eventos em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="cd19c-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="cd19c-104">O elemento `Event` no momento só tem suporte pelo Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="cd19c-104">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="cd19c-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="cd19c-105">Attributes</span></span>

|  <span data-ttu-id="cd19c-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="cd19c-106">Attribute</span></span>  |  <span data-ttu-id="cd19c-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="cd19c-107">Required</span></span>  |  <span data-ttu-id="cd19c-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd19c-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="cd19c-109">Type</span><span class="sxs-lookup"><span data-stu-id="cd19c-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="cd19c-110">Sim</span><span class="sxs-lookup"><span data-stu-id="cd19c-110">Yes</span></span>  | <span data-ttu-id="cd19c-111">Especifica o evento a ser manipulado.</span><span class="sxs-lookup"><span data-stu-id="cd19c-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="cd19c-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="cd19c-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="cd19c-113">Sim</span><span class="sxs-lookup"><span data-stu-id="cd19c-113">Yes</span></span>  | <span data-ttu-id="cd19c-p101">Especifica o estilo de execução para o manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="cd19c-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="cd19c-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="cd19c-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="cd19c-117">Sim</span><span class="sxs-lookup"><span data-stu-id="cd19c-117">Yes</span></span>  | <span data-ttu-id="cd19c-118">Especifica o nome da função para o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="cd19c-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="cd19c-119">Atributo de tipo</span><span class="sxs-lookup"><span data-stu-id="cd19c-119">Type attribute</span></span>

<span data-ttu-id="cd19c-p102">Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="cd19c-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="cd19c-123">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="cd19c-123">Event type</span></span>  |  <span data-ttu-id="cd19c-124">Descrição</span><span class="sxs-lookup"><span data-stu-id="cd19c-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="cd19c-125">O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.</span><span class="sxs-lookup"><span data-stu-id="cd19c-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="cd19c-126">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="cd19c-126">FunctionExecution attribute</span></span>

<span data-ttu-id="cd19c-p103">Obrigatório. DEVE ser definido como `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="cd19c-p103">Required. MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="cd19c-129">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="cd19c-129">FunctionName attribute</span></span>

<span data-ttu-id="cd19c-p104">Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cd19c-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```