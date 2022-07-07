---
title: Trabalhar com substituições estendidas do manifesto
description: Saiba como configurar recursos de extensibilidade com substituições estendidas do manifesto.
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43e9820f54f2812130f7f86529c52b20b92811a0
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659945"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Trabalhar com substituições estendidas do manifesto

Alguns recursos de extensibilidade dos Suplementos do Office são configurados com arquivos JSON hospedados em seu servidor, em vez de com o manifesto XML do suplemento.

> [!NOTE]
> Este artigo pressupõe que você esteja familiarizado com manifestos de suplementos do Office e sua função em suplementos. Leia o [manifesto XML dos Suplementos do Office](add-in-manifests.md), caso não tenha lido recentemente.

A tabela a seguir especifica os recursos de extensibilidade que exigem uma substituição estendida, juntamente com links para a documentação do recurso.

| Recurso | Instruções de desenvolvimento |
| :----- | :----- |
| Atalhos de teclado | [Adicionar atalhos de teclado personalizados aos suplementos do Office](../design/keyboard-shortcuts.md) |

O esquema que define o formato JSON é [o esquema de manifesto estendido](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!TIP]
> Este artigo é um pouco abstrato. Considere ler um dos artigos na tabela para adicionar clareza aos conceitos.

## <a name="tell-office-where-to-find-the-json-file"></a>Informe ao Office onde encontrar o arquivo JSON

Use o manifesto para informar ao Office onde encontrar o arquivo JSON. Imediatamente *abaixo* (não dentro) do **\<VersionOverrides\>** elemento no manifesto, adicione um [elemento ExtendedOverrides](/javascript/api/manifest/extendedoverrides) . Defina `Url` o atributo como a URL completa de um arquivo JSON. A seguir está um exemplo do elemento mais **\<ExtendedOverrides\>** simples possível.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

A seguir está um exemplo de um arquivo JSON de substituições estendidas muito simples. Ele atribui o atalho de teclado CTRL+SHIFT+A a uma função (definida em outro lugar) que abre o painel de tarefas do suplemento.

```json
{
    "actions": [
        {
            "id": "SHOWTASKPANE",
            "type": "ExecuteFunction",
            "name": "Show task pane for add-in"
        }
    ],
    "shortcuts": [
        {
            "action": "SHOWTASKPANE",
            "key": {
                "default": "CTRL+SHIFT+A"
            }
        }
    ]
}
```

## <a name="localize-the-extended-overrides-file"></a>Localizar o arquivo de substituições estendidas

Se o suplemento der suporte a várias localidades, `ResourceUrl` **\<ExtendedOverrides\>** você poderá usar o atributo do elemento para apontar o Office para um arquivo de recursos localizados. Apresentamos um exemplo a seguir.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Para obter mais detalhes sobre como criar e usar o arquivo de recursos, como fazer referência a seus recursos no arquivo de substituições estendidas e para obter opções adicionais não discutidas aqui, consulte Localizar substituições [estendidas](localization.md#localize-extended-overrides).
