---
title: Elemento Event no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51bbcd5a3d5abe60b850e88e4063e6bbc2da37bc
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450587"
---
# <a name="event-element"></a><span data-ttu-id="b0d17-102">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="b0d17-102">Event element</span></span>

<span data-ttu-id="b0d17-103">Define um manipulador de eventos em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0d17-103">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="b0d17-104">No `Event` momento, o elemento só tem suporte pelo Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="b0d17-104">The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="b0d17-105">Atributos</span><span class="sxs-lookup"><span data-stu-id="b0d17-105">Attributes</span></span>

|  <span data-ttu-id="b0d17-106">Atributo</span><span class="sxs-lookup"><span data-stu-id="b0d17-106">Attribute</span></span>  |  <span data-ttu-id="b0d17-107">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b0d17-107">Required</span></span>  |  <span data-ttu-id="b0d17-108">Descrição</span><span class="sxs-lookup"><span data-stu-id="b0d17-108">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="b0d17-109">Type</span><span class="sxs-lookup"><span data-stu-id="b0d17-109">Type</span></span>](#type-attribute)  |  <span data-ttu-id="b0d17-110">Sim</span><span class="sxs-lookup"><span data-stu-id="b0d17-110">Yes</span></span>  | <span data-ttu-id="b0d17-111">Especifica o evento a ser manipulado.</span><span class="sxs-lookup"><span data-stu-id="b0d17-111">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="b0d17-112">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="b0d17-112">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="b0d17-113">Sim</span><span class="sxs-lookup"><span data-stu-id="b0d17-113">Yes</span></span>  | <span data-ttu-id="b0d17-p101">Especifica o estilo de execução para o manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="b0d17-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="b0d17-116">FunctionName</span><span class="sxs-lookup"><span data-stu-id="b0d17-116">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="b0d17-117">Sim</span><span class="sxs-lookup"><span data-stu-id="b0d17-117">Yes</span></span>  | <span data-ttu-id="b0d17-118">Especifica o nome da função para o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="b0d17-118">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="b0d17-119">Atributo de tipo</span><span class="sxs-lookup"><span data-stu-id="b0d17-119">Type attribute</span></span>

<span data-ttu-id="b0d17-p102">Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="b0d17-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="b0d17-123">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="b0d17-123">Event type</span></span>  |  <span data-ttu-id="b0d17-124">Descrição</span><span class="sxs-lookup"><span data-stu-id="b0d17-124">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="b0d17-125">O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.</span><span class="sxs-lookup"><span data-stu-id="b0d17-125">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="b0d17-126">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="b0d17-126">FunctionExecution attribute</span></span>

<span data-ttu-id="b0d17-127">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="b0d17-127">Required.</span></span> <span data-ttu-id="b0d17-128">DEVE ser definido como `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="b0d17-128">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="b0d17-129">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="b0d17-129">FunctionName attribute</span></span>

<span data-ttu-id="b0d17-p104">Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0d17-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```
