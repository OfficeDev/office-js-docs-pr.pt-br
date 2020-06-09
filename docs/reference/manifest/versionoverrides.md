---
title: Elemento VersionOverrides no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para arquivos de manifesto de suplementos do Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cb23a78c336be891cdfa30262713ee3c80b9160f
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604493"
---
# <a name="versionoverrides-element"></a>Elemento VersionOverrides

O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sim  |  O namespace do esquema VersionOverrides. Os valores permitidos variam de acordo com o `<VersionOverrides>` valor **xsi: Type** do elemento e o valor **xsi: Type** do elemento pai `<OfficeApp>` . Consulte [namespace valores](#namespace-values) a seguir.|
|  **xsi:type**  |  Sim  | A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valores do namespace

A seguir, a lista o valor necessário do valor **xmlns** , dependendo do valor **xsi: Type** do elemento pai `<OfficeApp>` .

- **TaskPaneApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** oferece suporte somente à versão 1,0 do VersionOverrides e o **xmlns** deve ser `http://schemas.microsoft.com/office/contentappversionoverrides` .
- O **MailApp** suporta as versões 1,0 e 1,1 do VersionOverrides, portanto, o valor de **xmlns** varia de acordo com o `<VersionOverrides>` valor **xsi: Type** do elemento:
    - Quando **xsi: Type** é `VersionOverridesV1_0` , o **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Quando **xsi: Type** é `VersionOverridesV1_1` , o **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> Atualmente, somente o Outlook 2016 ou posterior oferece suporte ao esquema do VersionOverrides v 1.1 e ao `VersionOverridesV1_1` tipo.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Descrição**    |  Não   |  Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.|
|  **Requisitos**  |  Não   |  Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.|
|  [Hosts](hosts.md)                |  Sim  |  Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.  |
|  [Resources](resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  [EquivalentAddins](equivalentaddins.md)    |  Não  | Especifica os suplementos nativos (COM/XLL) equivalentes ao suplemento Web. O suplemento Web não será ativado se um suplemento nativo equivalente estiver instalado.|
|  **VersionOverrides**    |  Não  | Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do suplemento com emissores de token seguros, como o Azure Active Directory V 2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Não  |  Especifica uma coleção de permissões estendidas.<br><br>**Importante**: como a API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) está atualmente em versão prévia, os suplementos que usam o `ExtendedPermissions` elemento não podem ser publicados no AppSource ou implantados por meio da implantação centralizada. |

### <a name="versionoverrides-example"></a>Exemplo de VersionOverrides

Veja a seguir um exemplo de um `<VersionOverrides>` elemento típico, incluindo alguns elementos filhos que não são necessários, mas que são normalmente usados.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Implementar várias versões

Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.

Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior. O elemento filho `VersionOverrides` não herda os valores do elemento pai.

Para implementar o esquema do VersionOverrides v1.0 e do v1.1, o manifesto seria semelhante ao exemplo a seguir.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
