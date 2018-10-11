# <a name="getstarted-element"></a><span data-ttu-id="ea585-101">Elemento GetStarted</span><span class="sxs-lookup"><span data-stu-id="ea585-101">GetStarted element</span></span>

<span data-ttu-id="ea585-p101">Fornece informações usadas pelo texto explicativo que aparece quando o suplemento está instalado em hosts do Word, do Excel, do PowerPoint e do OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="ea585-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="ea585-104">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="ea585-104">Child elements</span></span>

| <span data-ttu-id="ea585-105">Elemento</span><span class="sxs-lookup"><span data-stu-id="ea585-105">Element</span></span>                       | <span data-ttu-id="ea585-106">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="ea585-106">Required</span></span> | <span data-ttu-id="ea585-107">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea585-107">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="ea585-108">Title</span><span class="sxs-lookup"><span data-stu-id="ea585-108">Title</span></span>](#title)               | <span data-ttu-id="ea585-109">Sim</span><span class="sxs-lookup"><span data-stu-id="ea585-109">Yes</span></span>      | <span data-ttu-id="ea585-110">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="ea585-110">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="ea585-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea585-111">Description</span></span>](#description)   | <span data-ttu-id="ea585-112">Sim</span><span class="sxs-lookup"><span data-stu-id="ea585-112">Yes</span></span>      | <span data-ttu-id="ea585-113">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ea585-113">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="ea585-114">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="ea585-114">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="ea585-115">Não</span><span class="sxs-lookup"><span data-stu-id="ea585-115">No</span></span>       | <span data-ttu-id="ea585-116">Uma URL para uma página que explica o suplemento em detalhes.</span><span class="sxs-lookup"><span data-stu-id="ea585-116">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="ea585-117">Title</span><span class="sxs-lookup"><span data-stu-id="ea585-117">Title</span></span> 

<span data-ttu-id="ea585-p102">Obrigatório. O título usado para o início do texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ea585-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="ea585-121">Descrição</span><span class="sxs-lookup"><span data-stu-id="ea585-121">Description</span></span>

<span data-ttu-id="ea585-p103">Obrigatório. A descrição / o conteúdo do corpo para o texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ea585-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="ea585-125">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="ea585-125">LearnMoreUrl</span></span>

<span data-ttu-id="ea585-p104">Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O atributo **resid** faz referência a uma identificação válida no elemento **Urls** na seção [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="ea585-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="ea585-129">**LearnMoreUrl** atualmente não é renderizado em clientes do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ea585-129">NOTE:**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="ea585-130">Recomendamos que você adicione essa URL para todos os clientes para que a URL seja processada quando ficar disponível.</span><span class="sxs-lookup"><span data-stu-id="ea585-130">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="ea585-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="ea585-131">See also</span></span>

<span data-ttu-id="ea585-132">Os exemplos de código a seguir utilizam o elemento **GetStarted** :</span><span class="sxs-lookup"><span data-stu-id="ea585-132">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="ea585-133">Suplemento web do Excel para manipular formatação de tabelas e gráficos</span><span class="sxs-lookup"><span data-stu-id="ea585-133">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="ea585-134">Suplemento do Word JavaScript SpecKit</span><span class="sxs-lookup"><span data-stu-id="ea585-134">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="ea585-135">Inserir gráficos do Excel usando o Microsoft Graph em um suplemento do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ea585-135">Insert Excel charts using Microsoft Graph in a PowerPoint Add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
