# <a name="event-element"></a><span data-ttu-id="352b5-101">Elemento Event</span><span class="sxs-lookup"><span data-stu-id="352b5-101">Event element</span></span>

<span data-ttu-id="352b5-102">Define um manipulador de eventos em um suplemento.</span><span class="sxs-lookup"><span data-stu-id="352b5-102">Defines an event handler in an add-in.</span></span>

> [!NOTE] 
> <span data-ttu-id="352b5-103">O elemento `Event` no momento só tem suporte no Outlook na Web no Office 365.</span><span class="sxs-lookup"><span data-stu-id="352b5-103">Note: The `Event` element is currently only supported by Outlook on the web in Office 365.</span></span>

## <a name="attributes"></a><span data-ttu-id="352b5-104">Atributos</span><span class="sxs-lookup"><span data-stu-id="352b5-104">Attributes</span></span>

|  <span data-ttu-id="352b5-105">Atributo</span><span class="sxs-lookup"><span data-stu-id="352b5-105">Attribute</span></span>  |  <span data-ttu-id="352b5-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="352b5-106">Required</span></span>  |  <span data-ttu-id="352b5-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="352b5-107">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="352b5-108">Tipo</span><span class="sxs-lookup"><span data-stu-id="352b5-108">Type</span></span>](#type-attribute)  |  <span data-ttu-id="352b5-109">Sim</span><span class="sxs-lookup"><span data-stu-id="352b5-109">Yes</span></span>  | <span data-ttu-id="352b5-110">Especifica o evento a ser manipulado.</span><span class="sxs-lookup"><span data-stu-id="352b5-110">Specifies the event to handle.</span></span> |
|  [<span data-ttu-id="352b5-111">FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="352b5-111">FunctionExecution</span></span>](#functionexecution-attribute)  |  <span data-ttu-id="352b5-112">Sim</span><span class="sxs-lookup"><span data-stu-id="352b5-112">Yes</span></span>  | <span data-ttu-id="352b5-p101">Especifica o estilo de execução do manipulador de eventos, assíncrono ou síncrono. No momento, somente os manipuladores de eventos síncronos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="352b5-p101">Specifies the execution style for the event handler, asynchronous or synchronous. Currently only synchronous event handlers are supported.</span></span> |
|  [<span data-ttu-id="352b5-115">FunctionName</span><span class="sxs-lookup"><span data-stu-id="352b5-115">FunctionName</span></span>](#functionname-attribute)  |  <span data-ttu-id="352b5-116">Sim</span><span class="sxs-lookup"><span data-stu-id="352b5-116">Yes</span></span>  | <span data-ttu-id="352b5-117">Especifica o nome da função para o manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="352b5-117">Specifies the function name for the event handler.</span></span> |

### <a name="type-attribute"></a><span data-ttu-id="352b5-118">Atributo de tipo</span><span class="sxs-lookup"><span data-stu-id="352b5-118">Type attribute</span></span>

<span data-ttu-id="352b5-p102">Obrigatório. Especifica quais eventos chamarão o manipulador de eventos. Os valores possíveis para este atributo são especificados na tabela a seguir.</span><span class="sxs-lookup"><span data-stu-id="352b5-p102">Required. Specifies which event will invoke the event handler. The possible values for this attribute are specified in the following table.</span></span>

|  <span data-ttu-id="352b5-122">Tipo de evento</span><span class="sxs-lookup"><span data-stu-id="352b5-122">Event type</span></span>  |  <span data-ttu-id="352b5-123">Descrição</span><span class="sxs-lookup"><span data-stu-id="352b5-123">Description</span></span>  |
|:-----|:-----|
|  `ItemSend`  |  <span data-ttu-id="352b5-124">O manipulador de eventos será chamado quando o usuário enviar uma mensagem ou convite de reunião.</span><span class="sxs-lookup"><span data-stu-id="352b5-124">The event handler will be invoked when the user sends a message or meeting invitation.</span></span>  |

### <a name="functionexecution-attribute"></a><span data-ttu-id="352b5-125">Atributo FunctionExecution</span><span class="sxs-lookup"><span data-stu-id="352b5-125">FunctionExecution attribute</span></span>

<span data-ttu-id="352b5-126">Obrigatório.</span><span class="sxs-lookup"><span data-stu-id="352b5-126">Required.</span></span> <span data-ttu-id="352b5-127">DEVE ser definido como `synchronous`.</span><span class="sxs-lookup"><span data-stu-id="352b5-127">MUST be set to `synchronous`.</span></span>

### <a name="functionname-attribute"></a><span data-ttu-id="352b5-128">Atributo FunctionName</span><span class="sxs-lookup"><span data-stu-id="352b5-128">FunctionName attribute</span></span>

<span data-ttu-id="352b5-p104">Obrigatório. Especifica o nome da função do manipulador de eventos. Esse valor deve coincidir com um nome de função no [arquivo de função](functionfile.md) do suplemento.</span><span class="sxs-lookup"><span data-stu-id="352b5-p104">Required. Specifies the function name of the event handler. This value must match a function name in the add-in's [function file](functionfile.md).</span></span>

```xml
<Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="itemSendHandler" /> 
```