---
title: Elemento VersionOverrides 1.1 no arquivo de manifesto para um complemento de email
description: Documentação de referência do elemento VersionOverrides 1.1 (email) para Office arquivos XML (manifesto de complementos).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d3187b1c6c60db47e23709f21f264d3c3b0538e4
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042162"
---
# <a name="versionoverrides-11-element-in-the-manifest-file-for-a-mail-add-in"></a>Elemento VersionOverrides 1.1 no arquivo de manifesto para um complemento de email

Esse elemento contém informações para recursos que não são suportados no manifesto base.

> [!NOTE]
> Este artigo pressupo que você está familiarizado com a visão geral do elemento [VersionOverrides](versionoverrides.md), que contém informações importantes sobre os atributos e variações do elemento.

**Tipo de suplemento:** Email

**Válido somente nestes esquemas VersionOverrides:**

- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos:**

- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)
- Alguns elementos filho podem estar associados a conjuntos de requisitos adicionais.

## <a name="child-elements"></a>Elementos filho

A tabela a seguir só se aplica à versão 1.1 dos elementos **VersionOverrides** e somente a complementos de email.

> [!NOTE]
> No iOS, `<WebApplicationInfo>` há suporte apenas. Todos os outros elementos filho **de VersionOverrides** são ignorados.

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Descrição](#description)    |  Não   |  Descreve o suplemento. |
|  [Requisitos](requirements.md)  |  Não   |  Especifica os conjuntos mínimos de requisitos que devem ser suportados para que a marcação no pai `VersionOverrides` entre em vigor. Isso sempre deve ser *mais* restritivo do `Requirements` que o elemento na parte base do manifesto.|
|  [Hosts](hosts.md)                |  Sim  |  Especifica uma coleção de Office aplicativos. O elemento Hosts filho substitui o elemento Hosts na parte pai do manifesto.  |
|  [Resources](resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  [EquivalentAddins](equivalentaddins.md)    |  Não  | Especifica os complementos nativos (COM/XLL) que são equivalentes ao complemento da Web. O complemento da Web não será ativado se um complemento nativo equivalente estiver instalado.|
|  **VersionOverrides**    |  Não  | Atualmente não é possível ser usável no VersionOverrides 1.1 para os complementos de email. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do complemento com emissores de token seguro, como Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Não  |  Especifica uma coleção de permissões estendidas. |

### <a name="description"></a>Descrição

Descreve o suplemento. Isso substitui o elemento `Description` em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](resources.md). O atributo do elemento Description não pode ter mais de 32 caracteres e é definido como o valor do atributo do elemento `resid` que contém o  `id` `String` texto.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nestes esquemas VersionOverrides:**

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos:**

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o pai `<VersionOverrides>` é tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o pai `<VersionOverrides>` é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o pai `<VersionOverrides>` é o tipo Mail 1.1.

## <a name="example"></a>Exemplo

Apresentamos um exemplo simples a seguir. Para exemplos mais completos, consulte os manifestos dos complementos de exemplo em Office exemplos de código [de complemento.](https://github.com/OfficeDev/PnP-OfficeAddins)

A seguir está um exemplo de um elemento típico, incluindo alguns elementos filho que não são `<VersionOverrides>` necessários, mas são normalmente usados.

```xml
<OfficeApp ... xsi:type="MailApp">
...
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
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Implementar várias versões

Um manifesto pode implementar várias versões do elemento `VersionOverrides` que é compatível com várias versões do esquema VersionOverrides. Isso pode ser feito para fornecer suporte opcional a novos recursos em um esquema mais recente, sem deixar de fornecer suporte a clientes antigos que não têm suporte para os novos recursos.

Para implementar várias versões, o elemento `VersionOverrides` da versão mais recente deve ser um filho do elemento `VersionOverrides` da versão anterior. O elemento filho `VersionOverrides` não herda os valores do elemento pai.

Para implementar o esquema VersionOverrides v1.0 e v1.1, o manifesto seria semelhante ao exemplo a seguir.

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
