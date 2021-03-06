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
# <a name="work-with-extended-overrides-of-the-manifest"></a><span data-ttu-id="f0288-103">Trabalhar com substituições estendidas do manifesto</span><span class="sxs-lookup"><span data-stu-id="f0288-103">Work with Extended Overrides of the manifest</span></span>

<span data-ttu-id="f0288-104">Alguns recursos de extensibilidade dos Complementos do Office são configurados com arquivos JSON hospedados em seu servidor, em vez de com o manifesto XML do complemento.</span><span class="sxs-lookup"><span data-stu-id="f0288-104">Some extensibility features of Office Add-ins are configured with JSON files that are hosted on your server, instead of with the add-in's XML manifest.</span></span>

> [!NOTE]
> <span data-ttu-id="f0288-105">Este artigo supõe que você esteja familiarizado com manifestos de complementos do Office e sua função em complementos. Leia o [manifesto XML de Complementos do Office](add-in-manifests.md), caso não tenha lido recentemente.</span><span class="sxs-lookup"><span data-stu-id="f0288-105">This article assumes that you're familiar with Office add-in manifests and their role in add-ins. Please read [Office Add-ins XML manifest](add-in-manifests.md), if you haven't recently.</span></span>

<span data-ttu-id="f0288-106">A tabela a seguir especifica os recursos de extensibilidade que exigem uma substituição estendida juntamente com links para a documentação do recurso.</span><span class="sxs-lookup"><span data-stu-id="f0288-106">The following table specifies the extensibility features that require an extended override along with links to documentation of the feature.</span></span>

| <span data-ttu-id="f0288-107">Recurso</span><span class="sxs-lookup"><span data-stu-id="f0288-107">Feature</span></span> | <span data-ttu-id="f0288-108">Instruções de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="f0288-108">Development Instructions</span></span> |
| :----- | :----- |
| <span data-ttu-id="f0288-109">Atalhos de teclado</span><span class="sxs-lookup"><span data-stu-id="f0288-109">Keyboard shortcuts</span></span> | [<span data-ttu-id="f0288-110">Adicionar atalhos de teclado personalizados aos seus Complementos do Office</span><span class="sxs-lookup"><span data-stu-id="f0288-110">Add Custom keyboard shortcuts to your Office Add-ins</span></span>](../design/keyboard-shortcuts.md) |

<span data-ttu-id="f0288-111">O esquema que define o formato JSON é [o esquema de manifesto estendido](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span><span class="sxs-lookup"><span data-stu-id="f0288-111">The schema that defines the JSON format is [extended-manifest schema](https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).</span></span>

> [!TIP]
> <span data-ttu-id="f0288-112">Este artigo é um pouco abstrato.</span><span class="sxs-lookup"><span data-stu-id="f0288-112">This article is somewhat abstract.</span></span> <span data-ttu-id="f0288-113">Considere ler um dos artigos na tabela para adicionar clareza aos conceitos.</span><span class="sxs-lookup"><span data-stu-id="f0288-113">Consider reading one of the articles in the table to add clarity to the concepts.</span></span>

## <a name="tell-office-where-to-find-the-json-file"></a><span data-ttu-id="f0288-114">Diga ao Office onde encontrar o arquivo JSON</span><span class="sxs-lookup"><span data-stu-id="f0288-114">Tell Office where to find the JSON file</span></span>

<span data-ttu-id="f0288-115">Use o manifesto para dizer ao Office onde encontrar o arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f0288-115">Use the manifest to tell Office where to find the JSON file.</span></span> <span data-ttu-id="f0288-116">Imediatamente *abaixo* (não dentro) `<VersionOverrides>` do elemento no manifesto, adicione um elemento [ExtendedOverrides.](../reference/manifest/extendedoverrides.md)</span><span class="sxs-lookup"><span data-stu-id="f0288-116">Immediately *below* (not inside) the `<VersionOverrides>` element in the manifest, add an [ExtendedOverrides](../reference/manifest/extendedoverrides.md) element.</span></span> <span data-ttu-id="f0288-117">De definir `Url` o atributo como a URL completa de um arquivo JSON.</span><span class="sxs-lookup"><span data-stu-id="f0288-117">Set the `Url` attribute to the full URL of a JSON file.</span></span> <span data-ttu-id="f0288-118">A seguir, um exemplo do elemento mais `<ExtendedOverrides>` simples possível.</span><span class="sxs-lookup"><span data-stu-id="f0288-118">The following is an example of the simplest possible `<ExtendedOverrides>` element.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json"></ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="f0288-119">A seguir, um exemplo de um arquivo JSON estendido muito simples substitui.</span><span class="sxs-lookup"><span data-stu-id="f0288-119">The following is an example of a very simple extended overrides JSON file.</span></span> <span data-ttu-id="f0288-120">Ele atribui o atalho de teclado CTRL+SHIFT+A a uma função (definida em outro lugar) que abre o painel de tarefas do complemento.</span><span class="sxs-lookup"><span data-stu-id="f0288-120">It assigns keyboard shortcut CTRL+SHIFT+A to a function (defined elsewhere) that opens the add-in's task pane.</span></span>

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

## <a name="localize-the-extended-overrides-file"></a><span data-ttu-id="f0288-121">Localize o arquivo de substituições estendidas</span><span class="sxs-lookup"><span data-stu-id="f0288-121">Localize the extended overrides file</span></span>

<span data-ttu-id="f0288-122">Se o seu add-in dá suporte a várias localidades, você pode usar o atributo do elemento para apontar o `ResourceUrl` Office para um arquivo de recursos `<ExtendedOverrides>` localizados.</span><span class="sxs-lookup"><span data-stu-id="f0288-122">If your add-in supports multiple locales, you can use the `ResourceUrl` attribute of the `<ExtendedOverrides>` element to point Office to a file of localized resources.</span></span> <span data-ttu-id="f0288-123">Apresentamos um exemplo a seguir.</span><span class="sxs-lookup"><span data-stu-id="f0288-123">The following is an example.</span></span>

```xml
    ...
    </VersionOverrides>  
    <ExtendedOverrides Url="https://contoso.com/addin/extended-overrides.json" 
                       ResourceUrl="https://contoso.com/addin/my-resources.json">
    </ExtendedOverrides>
</OfficeApp>
```

<span data-ttu-id="f0288-124">Para obter mais detalhes sobre como criar e usar o arquivo de recursos, como fazer referência a seus recursos no arquivo de substituições estendidas e para opções adicionais não discutidas aqui, consulte [Localize extended overrides](localization.md#localize-extended-overrides).</span><span class="sxs-lookup"><span data-stu-id="f0288-124">For more details about how to create and use the resources file, how to refer to its resources in the extended overrides file, and for additional options not discussed here, see [Localize extended overrides](localization.md#localize-extended-overrides).</span></span>
