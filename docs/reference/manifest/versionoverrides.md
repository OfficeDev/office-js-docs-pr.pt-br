---
title: Elemento VersionOverrides no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para arquivos XML (Add-ins manifest) Office.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555147"
---
# <a name="versionoverrides-element"></a>Elemento VersionOverrides

O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sim  |  O espaço de nome do esquema VersionOverrides. Os valores permitidos variam dependendo `<VersionOverrides>` do valor **xsi:tipo** deste elemento e do valor **xsi:tipo** do `<OfficeApp>` elemento pai. Veja [os valores do Namespace](#namespace-values) abaixo.|
|  **xsi:type**  |  Sim  | A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valores do namespace

A seguir, lista o valor necessário do valor **xmlns,** dependendo do valor **xsi:tipo** do `<OfficeApp>` elemento pai.

- **TaskPaneApp** suporta apenas a versão 1.0 do VersionOverrides, e os **xmlns** devem ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** suporta apenas a versão 1.0 do VersionOverrides, e os **xmlns** devem ser `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **O MailApp** suporta as versões 1.0 e 1.1 do VersionOverrides, de modo que o valor dos **xmlns** varia dependendo `<VersionOverrides>` do valor **xsi xsi:tipo** deste elemento:
    - Quando **xsi:type** é `VersionOverridesV1_0` , **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Quando **xsi:type** é `VersionOverridesV1_1` , **xmlns** deve ser `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> Atualmente, apenas Outlook 2016 ou posterior suporta o esquema VersionOverrides v1.1 e o `VersionOverridesV1_1` tipo.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Descrição**    |  Não   |  Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](resources.md). O `resid` atributo do elemento **Descrição** não pode ter mais de 32 caracteres e é definido para o valor do atributo do elemento que contém `id` o `String` texto.|
|  **Requisitos**  |  Não   |  Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.|
|  [Hosts](hosts.md)                |  Sim  |  Especifica uma coleção de Office aplicativos. O elemento hospedeiros da criança substitui o elemento Hosts na parte dos pais do manifesto.  |
|  [Resources](resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  [EquivalentAddins](equivalentaddins.md)    |  Não  | Especifica os complementos nativos (COM/XLL) que são equivalentes ao complemento da web. O complemento da web não será ativado se um complemento nativo equivalente for instalado.|
|  **VersionOverrides**    |  Não  | Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions). |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do complemento com emissores de tokens seguros, como Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Não  |  Especifica uma coleção de permissões estendidas. |

### <a name="versionoverrides-example"></a>Exemplo de VersionOverrides

A seguir, um exemplo de um `<VersionOverrides>` elemento típico, incluindo alguns elementos infantis que não são necessários, mas são tipicamente usados.

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
