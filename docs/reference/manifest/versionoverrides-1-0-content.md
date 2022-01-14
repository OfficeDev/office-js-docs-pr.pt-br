---
title: Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento de conteúdo
description: Documentação de referência do elemento VersionOverrides (conteúdo) para Office arquivos XML (manifesto de complementos).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 2a9cd431f0e8fb4a7abe49103522e04900d9bcfd
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042159"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-content-add-in"></a>Elemento VersionOverrides 1.0 no arquivo de manifesto para um complemento de conteúdo

Esse elemento contém informações para recursos que não são suportados no manifesto base.

> [!NOTE]
> Este artigo pressupo que você está familiarizado com a visão geral do elemento [VersionOverrides](versionoverrides.md), que contém informações importantes sobre os atributos e variações do elemento.

## <a name="child-elements"></a>Elementos filho

A tabela a seguir só se aplica à versão 1.0 dos elementos **VersionOverrides** e somente a complementos de conteúdo.

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **VersionOverrides**    |  Não  | Atualmente, não é possível ser usável no VersionOverrides 1.0 para os complementos de conteúdo. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Não  | Especifica detalhes sobre o registro do complemento com emissores de token seguro, como Azure Active Directory V2.0. |

## <a name="example"></a>Exemplo

Apresentamos um exemplo simples a seguir. Para exemplos mais completos, consulte os manifestos dos complementos de exemplo em Office exemplos de código [de complemento.](https://github.com/OfficeDev/PnP-OfficeAddins)

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
