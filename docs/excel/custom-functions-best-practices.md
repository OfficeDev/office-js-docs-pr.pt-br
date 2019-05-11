---
ms.date: 05/08/2019
description: Saiba mais sobre as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.
title: Práticas recomendadas para funções personalizadas
localization_priority: Normal
ms.openlocfilehash: d825f5a9f14e240ca5af3c3325cb646248d99ca9
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952100"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="ebfd9-103">Práticas recomendadas para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebfd9-103">Custom functions best practices</span></span>

<span data-ttu-id="ebfd9-104">Este artigo descreve as práticas recomendadas para o desenvolvimento de funções personalizadas para Excel.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="ebfd9-105">Associar os nomes de função com metadados JSON</span><span class="sxs-lookup"><span data-stu-id="ebfd9-105">Associating function names with JSON metadata</span></span>

<span data-ttu-id="ebfd9-106">Conforme descrito no artigo [visão geral de funções personalizados](custom-functions-overview.md), um projeto de funções personalizados deve incluir um arquivo JSON de metadados e um arquivo de script (JavaScript ou TypeScript) para formar uma função completa.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-106">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="ebfd9-107">Se você estiver usando `yo office` os metadados JSON podem ser gerados a partir dos comentários de código.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-107">If you are using `yo office` the JSON metadata can be generated from the code comments.</span></span> <span data-ttu-id="ebfd9-108">Caso contrário, você precisará criar o arquivo de metadados JSON manualmente.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-108">Otherwise you need to build the JSON metadata file manually.</span></span>

<span data-ttu-id="ebfd9-109">Para que uma função funcione corretamente, você precisa associar a propriedade da `id` função à implementação do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-109">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="ebfd9-110">Verifique se há uma associação, caso contrário, a função não será chamada.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-110">Make sure there is an association, otherwise the function will not be called.</span></span> <span data-ttu-id="ebfd9-111">O exemplo de código a seguir mostra como fazer a Associação usando `CustomFunctions.associate()` o método.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-111">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="ebfd9-112">A amostra define a função personalizada `add` e associa com o objeto no arquivo de metadados JSON onde o valor da `id` propriedade é **adicionar**.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-112">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="ebfd9-113">O JSON a seguir mostra os metadados JSON que estão associados ao código JavaScript da função personalizada anterior.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-113">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


<span data-ttu-id="ebfd9-114">Lembre-se das seguintes práticas recomendadas quando criar funções personalizadas no arquivo JavaScript e especificar as informações correspondentes no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-114">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="ebfd9-115">No arquivo de metadados JSON, verifique se o valor de cada propriedade `id` contém apenas caracteres alfanuméricos e pontos.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-115">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="ebfd9-116">No arquivo de metadados JSON, garanta que o valor de cada propriedade `id` seja exclusivo dentro do escopo do arquivo.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-116">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="ebfd9-117">Ou seja, nenhum objeto de duas funções no arquivo de metadados deve ter o mesmo valor `id`.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-117">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

* <span data-ttu-id="ebfd9-118">Não altere o valor de uma propriedade `id` no arquivo de metadados JSON, depois de mapeá-lo para um nome de função JavaScript correspondente.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-118">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="ebfd9-119">Para alterar o nome da função que os usuários finais visualizam no Excel, atualize a propriedade `name` no arquivo de metadados JSON. No entanto, nunca altere o valor de uma propriedade `id` depois de estabelecida.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-119">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="ebfd9-120">No arquivo JavaScript, especifique uma associação de função personalizada usando `CustomFunctions.associate` após cada função.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-120">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="ebfd9-121">O exemplo a seguir mostra os metadados JSON que correspondem às funções definidas nesse exemplo de código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-121">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="ebfd9-122">Os `id` valores `name` de propriedade e estão em letras maiúsculas, o que é uma prática recomendada ao descrever suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-122">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="ebfd9-123">Você só precisará adicionar esse JSON se estiver preparando seu próprio arquivo JSON manualmente e não usando a autogeração.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-123">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="ebfd9-124">Para obter mais informações sobre a autogeração, consulte [criar metadados JSON para funções personalizadas](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="ebfd9-124">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="additional-considerations"></a><span data-ttu-id="ebfd9-125">Considerações adicionais</span><span class="sxs-lookup"><span data-stu-id="ebfd9-125">Additional considerations</span></span>

<span data-ttu-id="ebfd9-126">Evite acessar o modelo de objeto de documento (DOM) direta ou indiretamente (por exemplo, usando jQuery) de sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-126">Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function.</span></span> <span data-ttu-id="ebfd9-127">No Excel no Windows, onde as funções personalizadas usam o [tempo de execução do JavaScript](custom-functions-runtime.md), as funções personalizadas não podem acessar o dom.</span><span class="sxs-lookup"><span data-stu-id="ebfd9-127">In Excel on Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ebfd9-128">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ebfd9-128">Next steps</span></span>
<span data-ttu-id="ebfd9-129">Saiba como [realizar solicitações da Web com funções personalizadas](custom-functions-web-reqs.md).</span><span class="sxs-lookup"><span data-stu-id="ebfd9-129">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ebfd9-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="ebfd9-130">See also</span></span>

* [<span data-ttu-id="ebfd9-131">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebfd9-131">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ebfd9-132">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ebfd9-132">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ebfd9-133">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="ebfd9-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
