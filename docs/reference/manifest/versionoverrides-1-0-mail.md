---
title: Elemento VersionOverrides 1.0 no arquivo de manifesto de um complemento de email
description: Documentação de referência do elemento VersionOverrides (email) para Office arquivos XML (manifesto de complementos).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5288c085c94ff6fc8ab8fc31711c5c8fa142e946
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340670"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>Elemento VersionOverrides 1.0 no arquivo de manifesto de um complemento de email

Esse elemento contém informações para recursos que não são suportados no manifesto base.

> [!NOTE]
> Este artigo supõe que você esteja familiarizado com a visão geral do elemento [VersionOverrides](versionoverrides.md), que contém informações importantes sobre os atributos e variações do elemento.

**Tipo de suplemento:** Email

**Válido somente nesses esquemas VersionOverrides**:

- Email 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- Alguns elementos filho podem estar associados a conjuntos de requisitos adicionais.

## <a name="child-elements"></a>Elementos filho

A tabela a seguir só se aplica à versão 1.0 dos elementos **VersionOverrides** e somente a complementos de email.

> [!NOTE]
> No iOS, há suporte apenas **para WebApplicationInfo** . Todos os outros elementos filho **de VersionOverrides** são ignorados.

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Descrição](#description)    |  Não   |  Descreve o suplemento. |
|  [Requisitos](requirements.md)  |  Não   |  Especifica os conjuntos mínimos de requisitos que devem ser suportados para que a marcação no **VersionOverrides** pai entre em vigor. Isso sempre deve ser *mais restritivo* do que o elemento **Requirements** na parte base do manifesto.|
|  [Hosts](hosts.md)                |  Sim  |  Especifica uma coleção de Office aplicativos. O elemento filho **Hosts** substitui o elemento **Hosts** na parte pai do manifesto.  |
|  [Resources](resources.md)    |  Sim  | Define um conjunto de recursos (cadeias de caracteres, URLs e imagens) consultado por outros elementos do manifesto.|
|  **VersionOverrides**    |  Não  | Define comandos de suplemento em uma versão mais recente do esquema. Para saber mais, confira o tópico [Implementar várias versões](#implementing-multiple-versions). |
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

Um manifesto pode implementar várias versões do **elemento VersionOverrides** que suportam versões diferentes do esquema VersionOverrides. Isso pode ser feito para oferecer suporte opcional a novos recursos em um esquema mais novo e ainda dar suporte a clientes mais antigos que não suportam os novos recursos.

Para implementar várias versões, o **elemento VersionOverrides** para a versão mais recente deve ser um filho do elemento para a `VersionOverrides` versão mais antiga. O elemento **VersionOverrides** filho não herda nenhum valor do pai.

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
