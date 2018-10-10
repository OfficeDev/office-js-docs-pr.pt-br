# <a name="host-element"></a><span data-ttu-id="86636-101">Elemento Host</span><span class="sxs-lookup"><span data-stu-id="86636-101">Host element</span></span>

<span data-ttu-id="86636-102">Especifica um tipo de aplicativo individual do Office em que o suplemento deve ser ativado.</span><span class="sxs-lookup"><span data-stu-id="86636-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="86636-103">A sintaxe do elemento **Host** varia dependendo se o elemento está definido no [manifesto básico](#basic-manifest) ou no nó [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="86636-103">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="86636-104">No entanto, a funcionalidade é a mesma.</span><span class="sxs-lookup"><span data-stu-id="86636-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="86636-105">Manifesto básico</span><span class="sxs-lookup"><span data-stu-id="86636-105">Basic manifest</span></span>

<span data-ttu-id="86636-106">Quando definido no manifesto básico (abaixo de [OfficeApp](officeapp.md)), o tipo de host é determinado pelo atributo `Name`.</span><span class="sxs-lookup"><span data-stu-id="86636-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="86636-107">Atributos</span><span class="sxs-lookup"><span data-stu-id="86636-107">Attributes</span></span>

| <span data-ttu-id="86636-108">Atributo</span><span class="sxs-lookup"><span data-stu-id="86636-108">Attribute</span></span>     | <span data-ttu-id="86636-109">Tipo</span><span class="sxs-lookup"><span data-stu-id="86636-109">Type</span></span>   | <span data-ttu-id="86636-110">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86636-110">Required</span></span> | <span data-ttu-id="86636-111">Descrição</span><span class="sxs-lookup"><span data-stu-id="86636-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="86636-112">Name</span><span class="sxs-lookup"><span data-stu-id="86636-112">Name</span></span>](#name) | <span data-ttu-id="86636-113">string</span><span class="sxs-lookup"><span data-stu-id="86636-113">string</span></span> | <span data-ttu-id="86636-114">obrigatório</span><span class="sxs-lookup"><span data-stu-id="86636-114">required</span></span> | <span data-ttu-id="86636-115">O nome do tipo de aplicativo host do Office.</span><span class="sxs-lookup"><span data-stu-id="86636-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="86636-116">Name</span><span class="sxs-lookup"><span data-stu-id="86636-116">Name</span></span>
<span data-ttu-id="86636-p102">Especifica o tipo de Host destinado por esse suplemento. O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="86636-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="86636-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="86636-119">`Document` (Word)</span></span>
- <span data-ttu-id="86636-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="86636-120">`Database` (Access)</span></span>
- <span data-ttu-id="86636-121">`Mailbox` Outlook</span><span class="sxs-lookup"><span data-stu-id="86636-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="86636-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="86636-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="86636-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="86636-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="86636-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="86636-124">`Project` (Project)</span></span>
- <span data-ttu-id="86636-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="86636-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="86636-126">Exemplo</span><span class="sxs-lookup"><span data-stu-id="86636-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="86636-127">Nó VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="86636-127">VersionOverrides node</span></span>
<span data-ttu-id="86636-128">Quando definido em [VersionOverrides](versionoverrides.md), o tipo de host é determinado pelo atributo `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="86636-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="86636-129">Atributos</span><span class="sxs-lookup"><span data-stu-id="86636-129">Attributes</span></span>

|  <span data-ttu-id="86636-130">Atributo</span><span class="sxs-lookup"><span data-stu-id="86636-130">Attribute</span></span>  |  <span data-ttu-id="86636-131">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86636-131">Required</span></span>  |  <span data-ttu-id="86636-132">Descrição</span><span class="sxs-lookup"><span data-stu-id="86636-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="86636-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="86636-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="86636-134">Sim</span><span class="sxs-lookup"><span data-stu-id="86636-134">Yes</span></span>  | <span data-ttu-id="86636-135">Descreve o host do Office ao qual se aplicam essas configurações.</span><span class="sxs-lookup"><span data-stu-id="86636-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="86636-136">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="86636-136">Child elements</span></span>

|  <span data-ttu-id="86636-137">Elemento</span><span class="sxs-lookup"><span data-stu-id="86636-137">Element</span></span> |  <span data-ttu-id="86636-138">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="86636-138">Required</span></span>  |  <span data-ttu-id="86636-139">Descrição</span><span class="sxs-lookup"><span data-stu-id="86636-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="86636-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="86636-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="86636-141">Sim</span><span class="sxs-lookup"><span data-stu-id="86636-141">Yes</span></span>   |  <span data-ttu-id="86636-142">Define as configurações do fator forma da área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="86636-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="86636-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="86636-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="86636-144">Não</span><span class="sxs-lookup"><span data-stu-id="86636-144">No</span></span>   |  <span data-ttu-id="86636-p103">Define as configurações do fator forma móvel. **Observação:** esse elemento só tem suporte no Outlook para iOS.</span><span class="sxs-lookup"><span data-stu-id="86636-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="86636-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="86636-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="86636-148">Não</span><span class="sxs-lookup"><span data-stu-id="86636-148">No</span></span>   |  <span data-ttu-id="86636-149">Define as configurações de todos os fatores forma.</span><span class="sxs-lookup"><span data-stu-id="86636-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="86636-150">Usado somente pelas funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="86636-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="86636-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="86636-151">xsi:type</span></span>

<span data-ttu-id="86636-152">Controla a qual host do Office (Word, Excel, PowerPoint, Outlook, OneNote) as configurações contidas se aplicam.</span><span class="sxs-lookup"><span data-stu-id="86636-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="86636-153">O valor deve ser uma das seguintes opções:</span><span class="sxs-lookup"><span data-stu-id="86636-153">The value must be one of the following:</span></span>

- <span data-ttu-id="86636-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="86636-154">`Document` (Word)</span></span>
- <span data-ttu-id="86636-155">`MailHost` Outlook</span><span class="sxs-lookup"><span data-stu-id="86636-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="86636-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="86636-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="86636-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="86636-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="86636-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="86636-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="86636-159">Exemplo de host</span><span class="sxs-lookup"><span data-stu-id="86636-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
