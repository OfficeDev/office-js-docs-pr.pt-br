---
title: Elemento VersionOverrides no arquivo de manifesto
description: ''
ms.date: 01/29/2019
localization_priority: Normal
ms.openlocfilehash: 897c2203ef6ae84911b7f269ee8a2c88aec36bd0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452064"
---
# <a name="versionoverrides-element"></a>Elemento VersionOverrides

O elemento raiz que contém informações para os comandos de suplemento implementados pelo suplemento. **VersionOverrides** é um elemento filho do elemento [OfficeApp](./officeapp.md) no manifesto. Ele recebe suporte no esquema de manifesto v1.1 e posterior, mas é definido no esquema VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sim  |  O local do esquema, que deve ser `http://schemas.microsoft.com/office/mailappversionoverrides` quando `xsi:type` for `VersionOverridesV1_0` e `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` quando `xsi:type` for `VersionOverridesV1_1`.|
|  **xsi:type**  |  Sim  | A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`. |

> [!NOTE]
> Atualmente, somente o Outlook 2016 ou posterior oferece suporte ao esquema do VersionOverrides `VersionOverridesV1_1` v 1.1 e ao tipo.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **Descrição**    |  Não   |  Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](./resources.md). O atributo `resid` do elemento **Description** está definido como o valor do atributo `id` do elemento `String` que contém o texto.|
|  **Requisitos**  |  Não   |  Especifica o conjunto de requisitos mínimos e a versão do Office.js exigida pelo suplemento. Isso substitui o elemento `Requirements` na parte pai do manifesto.|
|  [Hosts](./hosts.md)                |  Sim  |  Especifica um conjunto de hosts do Office. O elemento filho Hosts substitui o elemento Hosts na parte pai do manifesto.  |
|  [Resources](./resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  **VersionOverrides**    |  Não  | Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions). |
|  **WebApplicationInfo**    |  Não  | Especifica detalhes sobre o aplicativo Web associado do suplemento. |

### <a name="versionoverrides-example"></a>Exemplo de VersionOverrides

Veja a seguir um exemplo de um elemento `<VersionOverrides>` típico, incluindo alguns elementos filhos que não são necessários, mas que são normalmente usados.

```xml
<OfficeApp>
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
<OfficeApp>
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
