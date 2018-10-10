# <a name="group-element"></a>Elemento Group

Define um grupo de controles de interface do usuário em uma guia. Em guias personalizadas, o suplemento pode criar até 10 grupos. Cada grupo é limitado a 6 controles, independentemente de qual guia aparece. Suplementos são limitados a uma guia personalizada.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [id](#id-attribute)  |  Sim  | Uma ID exclusiva do grupo.|

### <a name="id-attribute"></a>Atributo id

Obrigatório. O identificador exclusivo do grupo. É uma cadeia de caracteres com, no máximo, 125 caracteres. Esse valor deve ser exclusivo dentro o manifesto, ou o grupo não será processado.

## <a name="child-elements"></a>Elementos filho
|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Rótulo](#label)      | Sim |  O rótulo para a CustomTab ou um grupo.  |
|  [Control](#control)    | Sim |  Conjunto de um ou mais objetos Control.  |

### <a name="label"></a>Rótulo 

Obrigatório. O rótulo do grupo. O atributo **resid** deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** do elemento [Resources](resources.md).

### <a name="control"></a>Control
Um grupo exige pelo menos um controle.

```xml
<Group id="msgreadCustomTab.grp1">
    <Label resid="residCustomTabGroupLabel"/>
    <Control xsi:type="Button" id="Button2">
    <!-- information on the control -->
    </Control>
    <!-- other controls, as needed -->
</Group>
```