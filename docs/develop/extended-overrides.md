---
title: Trabalhar com substituições estendidas do manifesto
description: Saiba como configurar recursos de extensibilidade com substituições estendidas do manifesto.
ms.date: 02/23/2021
localization_priority: Normal
ms.openlocfilehash: 4eb8936e8a01b81a3883f848446d20ebf4ecf863
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505566"
---
# <a name="work-with-extended-overrides-of-the-manifest"></a>Trabalhar com substituições estendidas do manifesto

Alguns recursos de extensibilidade dos Complementos do Office são configurados com arquivos JSON hospedados em seu servidor, em vez de com o manifesto XML do complemento.

> [!NOTE]
> Este artigo supõe que você esteja familiarizado com manifestos de complementos do Office e sua função em complementos. Leia o [manifesto XML de Complementos do Office](add-in-manifests.md), caso não tenha lido recentemente.

A tabela a seguir especifica os recursos de extensibilidade que exigem uma substituição estendida juntamente com links para a documentação do recurso.

| Recurso | Instruções de desenvolvimento |
| :----- | :----- |
| Atalhos de teclado | [Adicionar atalhos de teclado personalizados aos seus Complementos do Office](../design/keyboard-shortcuts.md) |

O esquema que define o formato JSON é [o esquema de manifesto estendido](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).

> [!TIP]
> Este artigo é um pouco abstrato. Considere ler um dos artigos na tabela para adicionar clareza aos conceitos.

## <a name="tell-office-where-to-find-the-json-file"></a>Diga ao Office onde encontrar o arquivo JSON

Use o manifesto para dizer ao Office onde encontrar o arquivo JSON. Imediatamente *abaixo* (não dentro) `<VersionOverrides>` do elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md) De definir `Url` o atributo como a URL completa de um arquivo JSON. A seguir, um exemplo do elemento mais `<ExtendedOverrides>` simples possível.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

A seguir, um exemplo de um arquivo JSON estendido muito simples substitui. Ele atribui o atalho de teclado CTRL+SHIFT+A a uma função (definida em outro lugar) que abre o painel de tarefas do complemento.

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

## <a name="localize-the-extended-overrides-file"></a>Localize o arquivo de substituições estendidas

Se o seu add-in dá suporte a várias localidades, você pode usar o atributo do elemento para apontar o `ResourceUrl` Office para um arquivo de recursos `<ExtendedOverrides>` localizados. Apresentamos um exemplo a seguir.

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

Para obter mais detalhes sobre como criar e usar o arquivo de recursos, como fazer referência a seus recursos no arquivo de substituições estendidas e para opções adicionais não discutidas aqui, consulte [Localize extended overrides](localization.md#localize-extended-overrides).
