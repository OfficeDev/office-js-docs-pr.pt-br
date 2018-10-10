# <a name="mobileformfactor-element"></a><span data-ttu-id="b0db3-101">Elemento MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="b0db3-101">MobileFormFactor element</span></span>

<span data-ttu-id="b0db3-p101">Especifica as configurações de um suplemento para um fator forma móvel. Ele contém todas as informações do suplemento para o fator forma móvel, exceto para o nó **Resources**.</span><span class="sxs-lookup"><span data-stu-id="b0db3-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="b0db3-p102">Cada definição de **MobileFormFactor** contém o elemento **FunctionFile** e um ou mais elementos **ExtensionPoint**. Para saber mais, confira [Elemento FunctionFile](functionfile.md) e [Elemento ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="b0db3-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="b0db3-p103">O elemento **MobileFormFactor** é definido no esquema VersionOverrides 1.1. O elemento [VersionOverrides](versionoverrides.md) incluído deve ter um valor de atributo `xsi:type` de `VersionOverridesV1_1`.</span><span class="sxs-lookup"><span data-stu-id="b0db3-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b0db3-108">Elementos filhos</span><span class="sxs-lookup"><span data-stu-id="b0db3-108">Child elements</span></span>

| <span data-ttu-id="b0db3-109">Elemento</span><span class="sxs-lookup"><span data-stu-id="b0db3-109">Element</span></span>                               | <span data-ttu-id="b0db3-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="b0db3-110">Required</span></span> | <span data-ttu-id="b0db3-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="b0db3-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="b0db3-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b0db3-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="b0db3-113">Sim</span><span class="sxs-lookup"><span data-stu-id="b0db3-113">Yes</span></span>      | <span data-ttu-id="b0db3-114">Define onde um suplemento expõe a funcionalidade.</span><span class="sxs-lookup"><span data-stu-id="b0db3-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="b0db3-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="b0db3-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="b0db3-116">Sim</span><span class="sxs-lookup"><span data-stu-id="b0db3-116">Yes</span></span>      | <span data-ttu-id="b0db3-117">Uma URL para um arquivo que contém funções JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b0db3-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="b0db3-118">Exemplo de MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="b0db3-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
