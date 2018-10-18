# <a name="supportssharedfolders-element"></a><span data-ttu-id="df02b-101">Elemento SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="df02b-101">SupportsSharedFolders element</span></span>

<span data-ttu-id="df02b-102">Define se o suplemento do Outlook está disponível em cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="df02b-102">Defines whether the Outlook add-in is available in delegate scenarios and is set to false by default.</span></span> <span data-ttu-id="df02b-103">O elemento **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="df02b-103">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="df02b-104">É definido como *false* por padrão.</span><span class="sxs-lookup"><span data-stu-id="df02b-104">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="df02b-105">O elemento SupportsSharedFolders só está disponível no [Conjunto de Requerimentos em versão prévia para suplementos do Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) para o Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="df02b-105">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="df02b-106">Os suplementos que usam esse elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.</span><span class="sxs-lookup"><span data-stu-id="df02b-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="df02b-107">A seguir apresentamos um exemplo do elemento **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="df02b-107">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
