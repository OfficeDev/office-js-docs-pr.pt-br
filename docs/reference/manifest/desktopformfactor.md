# <a name="desktopformfactor-element"></a><span data-ttu-id="20186-101">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="20186-101">DesktopFormFactor element</span></span>

<span data-ttu-id="20186-p101">Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office para Windows, Office para Mac e Office Online. Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="20186-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="20186-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="20186-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="20186-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="20186-107">Child elements</span></span>

| <span data-ttu-id="20186-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="20186-108">Element</span></span>                               | <span data-ttu-id="20186-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="20186-109">Required</span></span> | <span data-ttu-id="20186-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="20186-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="20186-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="20186-111">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="20186-112">Sim</span><span class="sxs-lookup"><span data-stu-id="20186-112">Yes</span></span>      | <span data-ttu-id="20186-113">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="20186-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="20186-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="20186-114">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="20186-115">Sim</span><span class="sxs-lookup"><span data-stu-id="20186-115">Yes</span></span>      | <span data-ttu-id="20186-116">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="20186-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="20186-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="20186-117">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="20186-118">Não</span><span class="sxs-lookup"><span data-stu-id="20186-118">No</span></span>       | <span data-ttu-id="20186-119">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="20186-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="20186-120">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="20186-120">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="20186-121">Não</span><span class="sxs-lookup"><span data-stu-id="20186-121">No</span></span> | <span data-ttu-id="20186-122">Define se o suplemento do Outlook está disponível nos cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="20186-122">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="20186-123">**Importante**: esse elemento só está disponível no requisito de visualização do Outlook suplementos definido contra o Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="20186-123">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="20186-124">Suplementos que usam esse elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="20186-124">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="20186-125">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="20186-125">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
