# <a name="customtab-element"></a>Elemento CustomTab

Na faixa de opções, especifique qual guia e grupo para seus comandos de suplemento. Isso pode ser realizado na guia padrão (**Página inicial**, **Mensagem** ou **Reunião**) ou em uma guia personalizada definida pelo suplemento.

Nas guias personalizadas, o suplemento poderá criar até 10 grupos. Cada grupo está limitado a seis controles, independentemente da guia na qual aparece. Os suplementos estão limitados a uma guia personalizada.

O atributo **id** deve ser único dentro do manifesto.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Group](group.md)      | Sim |  Define um grupo de comandos  |
|  [Label](#label-tab)      | Sim |  O rótulo para CustomTab ou Group.  |
|  [Control](control.md)    | Sim |  Conjunto de um ou mais objetos Control.  |

### <a name="group"></a>Group

Obrigatório. Confira [Elemento Group](group.md).

### <a name="label-tab"></a>Label (Tab)

Obrigatório. O rótulo da guia personalizada. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).


## <a name="customtab-example"></a>Exemplo CustomTab

```xml
<ExtensionPoint xsi:type="MessageReadCommandSurface">
  <CustomTab id="TabCustom1">
    <Group id="msgreadCustomTab.grp1">
    </Group>
    <Label resid="customTabLabel1"/>
  </CustomTab>
</ExtensionPoint>
```