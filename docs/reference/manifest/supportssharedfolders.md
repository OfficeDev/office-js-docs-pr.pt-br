# <a name="supportssharedfolders-element"></a>Elemento SupportsSharedFolders

Define se o suplemento do Outlook está disponível em cenários de representante. O elemento **SupportsSharedFolders** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). É definido como *false* por padrão.

> [!IMPORTANT]
> O elemento SupportsSharedFolders só está disponível no [Conjunto de Requerimentos em versão prévia para suplementos do Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) para o Exchange Online. Os suplementos que usam esse elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada.

A seguir apresentamos um exemplo do elemento **SupportsSharedFolders** .

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
