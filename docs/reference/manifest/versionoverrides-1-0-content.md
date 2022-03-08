---
title: Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento de conteúdo
description: Documentação de referência do elemento VersionOverrides (conteúdo) Office arquivos XML (manifesto de complementos).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ef083ef5df322c230292625576e36db8923d00c
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341048"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento de conteúdo

Esse elemento contém informações para recursos que não são suportados no manifesto base.

> [!NOTE]
> Este artigo supõe que você esteja familiarizado com a visão geral do elemento [VersionOverrides](versionoverrides.md), que contém informações importantes sobre os atributos e variações do elemento.

## <a name="child-elements"></a>Elementos filho

A tabela a seguir só se aplica à versão 1.0 dos elementos **VersionOverrides** e somente a complementos de conteúdo.

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Não  | Atualmente, não é possível ser usável no VersionOverrides 1.0 para os complementos de conteúdo. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do complemento com emissores de token seguro, como Azure Active Directory V2.0. |

## <a name="example"></a>Exemplo

Apresentamos um exemplo simples a seguir. Para obter exemplos mais complexos, consulte os manifestos dos complementos de exemplo [em Office exemplos de código de complemento](https://github.com/OfficeDev/PnP-OfficeAddins).

```xml
<OfficeApp ... xsi:type="Content">
...
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/contentappversionoverrides" xsi:type="VersionOverridesV1_0">
        <WebApplicationInfo>
            <Id>$application_GUID here$</Id>
            <Resource>api://localhost:44355/$application_GUID here$</Resource>
            <Scopes>
                <Scope>Files.Read.All</Scope>
                <Scope>profile</Scope>
            </Scopes>
        </WebApplicationInfo>
    </VersionOverrides>
...
</OfficeApp>
```
