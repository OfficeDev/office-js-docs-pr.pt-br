# <a name="desktopformfactor-element"></a><span data-ttu-id="47826-101">Elemento DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="47826-101">DesktopFormFactor element</span></span>

<span data-ttu-id="47826-p101">Especifica as configurações de um suplemento para o fator forma da área de trabalho. O fator de forma da área de trabalho inclui o Office para Windows, Office para Mac e Office Online. Ele contém todas as informações do suplemento para o fator forma da área de trabalho, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="47826-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="47826-p102">Cada definição de DesktopFormFactor contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="47826-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="47826-107">O elemento SupportsSharedFolders só está disponível no Conjunto de Requerimentos em versão prévia para suplementos do Outlook contra o Exhange Online.</span><span class="sxs-lookup"><span data-stu-id="47826-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="47826-108">Suplementos que usam esse elemento não são permitidos na Office Store ou na Implantação Centralizada.</span><span class="sxs-lookup"><span data-stu-id="47826-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="47826-109">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="47826-109">Child elements</span></span>

| <span data-ttu-id="47826-110">Elemento</span><span class="sxs-lookup"><span data-stu-id="47826-110">Element</span></span>                               | <span data-ttu-id="47826-111">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="47826-111">Required</span></span> | <span data-ttu-id="47826-112">Descrição</span><span class="sxs-lookup"><span data-stu-id="47826-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="47826-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="47826-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="47826-114">Sim</span><span class="sxs-lookup"><span data-stu-id="47826-114">Yes</span></span>      | <span data-ttu-id="47826-115">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="47826-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="47826-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="47826-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="47826-117">Sim</span><span class="sxs-lookup"><span data-stu-id="47826-117">Yes</span></span>      | <span data-ttu-id="47826-118">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="47826-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="47826-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="47826-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="47826-120">Não</span><span class="sxs-lookup"><span data-stu-id="47826-120">No</span></span>       | <span data-ttu-id="47826-121">Define o texto explicativo que aparece ao instalar o suplemento em hosts do Word, Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="47826-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="47826-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="47826-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="47826-123">Não</span><span class="sxs-lookup"><span data-stu-id="47826-123">No</span></span>       | <span data-ttu-id="47826-124">Define se o suplemento do Outlook está disponível nos cenários de representante e é definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="47826-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="47826-125">Conjunto de requisitos em versão prévia</span><span class="sxs-lookup"><span data-stu-id="47826-125">Outlook add-in API Preview requirement set</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="47826-126">Exemplo de DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="47826-126">DesktopFormFactor example</span></span>

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
