---
title: Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento do painel de tarefas
description: Documentação de referência do elemento VersionOverrides (painel de tarefas) para arquivos XML (manifesto de Office de complementos).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f2d6867db8a8b35d4296b9907e4dbbb440ea28db
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340243"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-task-pane-add-in"></a>Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento do painel de tarefas

Esse elemento contém informações para recursos que não são suportados no manifesto base.

> [!NOTE]
> Este artigo supõe que você esteja familiarizado com a visão geral do elemento [VersionOverrides](versionoverrides.md), que contém informações importantes sobre os atributos e variações do elemento.

**Tipo de suplemento:** Painel de tarefas

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) (Obrigatório para Excel, PowerPoint e Word).)
- Alguns elementos filho podem estar associados a conjuntos de requisitos adicionais.

## <a name="child-elements"></a>Elementos filho

A tabela a seguir só se aplica à versão 1.0 dos elementos **VersionOverrides** e somente aos complementos do painel de tarefas.

> [!NOTE]
> No iOS, há suporte apenas **para WebApplicationInfo** . Todos os outros elementos filho **de VersionOverrides** são ignorados.

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Descrição](#description)    |  Não   |  Descreve o suplemento. |
|  [Requisitos](requirements.md)  |  Não   |  Especifica os conjuntos mínimos de requisitos que devem ser suportados para que a marcação no **VersionOverrides** pai entre em vigor. Isso sempre deve ser *mais restritivo* do que o elemento **Requirements** na parte base do manifesto.|
|  [Hosts](hosts.md)                |  Sim  |  Especifica uma coleção de Office aplicativos. O elemento Hosts filho substitui o elemento Hosts na parte pai do manifesto.  |
|  [Resources](resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  [EquivalentAddins](equivalentaddins.md)    |  Não  | Especifica os complementos nativos (COM/XLL) que são equivalentes ao complemento da Web. O complemento da Web não será ativado se um complemento nativo equivalente estiver instalado.|
|  **VersionOverrides**    |  Não  | Atualmente não é possível ser usável no VersionOverrides 1.0 para os complementos do taskpane. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do complemento com emissores de token seguro, como Azure Active Directory V2.0. |

### <a name="description"></a>Descrição

Descreve o suplemento. Isso substitui o elemento **Description** em qualquer parte pai do manifesto. O texto da descrição está contido em um elemento filho do elemento **LongString**, contido no elemento [Resources](resources.md). O `resid` atributo do elemento **Description** não pode ter mais de 32 `id` caracteres e deve corresponder ao valor do atributo de um elemento filho do **elemento ShortString** contido no elemento [Resources](resources.md) .

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="example"></a>Exemplo

Apresentamos um exemplo simples a seguir. Para obter exemplos mais complexos, consulte os manifestos dos complementos de exemplo [em Office exemplos de código de complemento](https://github.com/OfficeDev/PnP-OfficeAddins).

```xml
<OfficeApp ... xsi:type="Taskpane">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="Workbook">
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
