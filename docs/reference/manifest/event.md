---
title: Elemento Event no arquivo de manifesto
description: Define um manipulador de eventos em um suplemento.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02037a54ad4b7e91a3697b53b04fa30e8a4909a9
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718227"
---
# <a name="event-element"></a><span data-ttu-id="57c09-103">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="57c09-103">Event element</span></span>

<span data-ttu-id="57c09-104">Define um manipulador de eventos em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="57c09-104">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="57c09-105">No `Event` momento, o elemento só tem suporte pelo Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="57c09-105">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="57c09-106">Atributos</span><span class="sxs-lookup"><span data-stu-id="57c09-106">Attributes</span></span>

|  <span data-ttu-id="57c09-107">Atributo</span><span class="sxs-lookup"><span data-stu-id="57c09-107">Attribute</span></span>  |  <span data-ttu-id="57c09-108">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="57c09-108">Required</span></span>  |  <span data-ttu-id="57c09-109">Descrição</span><span class="sxs-lookup"><span data-stu-id="57c09-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="57c09-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="57c09-110">Type</span></span>](#type-attribute)  |  <span data-ttu-id="57c09-111">Sim</span><span class="sxs-lookup"><span data-stu-id="57c09-111">Yes</span></span>  | <span data-ttu-id="57c09-112">Especifica o evento a ser manipulado.</span><span class="sxs-lookup"><span data-stu-id="57c09-112">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="57c09-113">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="57c09-113">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="57c09-114">Sim</span><span class="sxs-lookup"><span data-stu-id="57c09-114">Yes</span></span>  | <span data-ttu-id="57c09-p101">Especifica o estilo de execução para o manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="57c09-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="57c09-117">FunctionName</span><span class="sxs-lookup"><span data-stu-id="57c09-117">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="57c09-118">Sim</span><span class="sxs-lookup"><span data-stu-id="57c09-118">Yes</span></span>  | <span data-ttu-id="57c09-119">Especifica o nome da função para o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="57c09-119">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="57c09-120">Atributo de tipo</span><span class="sxs-lookup"><span data-stu-id="57c09-120">Type attribute</span></span>

<span data-ttu-id="57c09-p102">Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="57c09-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="57c09-124">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="57c09-124">Event type</span></span>  |  <span data-ttu-id="57c09-125">Descrição</span><span class="sxs-lookup"><span data-stu-id="57c09-125">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="57c09-126">O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.</span><span class="sxs-lookup"><span data-stu-id="57c09-126">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="57c09-127">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="57c09-127">FunctionExecution attribute</span></span>

<span data-ttu-id="57c09-128">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="57c09-128">Required.</span></span> <span data-ttu-id="57c09-129">DEVE ser definido como `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="57c09-129">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="57c09-130">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="57c09-130">FunctionName attribute</span></span>

<span data-ttu-id="57c09-p104">Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="57c09-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
