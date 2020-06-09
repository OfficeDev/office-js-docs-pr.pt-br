---
ms.date: 04/29/2020
description: Localize suas funções personalizadas do Excel.
title: Localizar funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 427bff029c5e85caa216f628df450525ee187c17
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609293"
---
# <a name="localize-custom-functions"></a><span data-ttu-id="8b90c-103">Localizar funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8b90c-103">Localize custom functions</span></span>

<span data-ttu-id="8b90c-104">Você pode localizar o suplemento e seus nomes de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8b90c-104">You can localize both your add-in and your custom function names.</span></span> <span data-ttu-id="8b90c-105">Para fazer isso, forneça nomes de função localizados no arquivo JSON de funções e informações de localidade no arquivo de manifesto XML.</span><span class="sxs-lookup"><span data-stu-id="8b90c-105">To do so, provide localized function names in the functions' JSON file and locale information in the XML manifest file.</span></span>

>[!IMPORTANT]
> <span data-ttu-id="8b90c-106">Os metadados gerados automaticamente não funcionam para localização, portanto, você precisa atualizar o arquivo JSON manualmente.</span><span class="sxs-lookup"><span data-stu-id="8b90c-106">Auto-generated metadata doesn't work for localization so you need to update the JSON file manually.</span></span> <span data-ttu-id="8b90c-107">Para saber como fazer isso, confira [metadados de funções personalizadas no Excel](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="8b90c-107">To learn how to do this, see [Metadata for custom functions in Excel](custom-functions-json.md)</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a><span data-ttu-id="8b90c-108">Localizar nomes de função</span><span class="sxs-lookup"><span data-stu-id="8b90c-108">Localize function names</span></span>

<span data-ttu-id="8b90c-109">Para localizar suas funções personalizadas, crie um novo arquivo de metadados JSON para cada idioma.</span><span class="sxs-lookup"><span data-stu-id="8b90c-109">To localize your custom functions, create a new JSON metadata file for each language.</span></span> <span data-ttu-id="8b90c-110">Em cada arquivo JSON de idioma, crie `name` e `description` Propriedades no idioma de destino.</span><span class="sxs-lookup"><span data-stu-id="8b90c-110">In each language JSON file, create `name` and `description` properties in the target language.</span></span> <span data-ttu-id="8b90c-111">O arquivo padrão para inglês é chamado de **funções. JSON**.</span><span class="sxs-lookup"><span data-stu-id="8b90c-111">The default file for English is named **functions.json**.</span></span> <span data-ttu-id="8b90c-112">Use a localidade no nome do arquivo para cada arquivo JSON adicional, como **funções-de. JSON** para ajudá-lo a identificá-los.</span><span class="sxs-lookup"><span data-stu-id="8b90c-112">Use the locale in the filename for each additional JSON file, such as **functions-de.json** to help identify them.</span></span>

<span data-ttu-id="8b90c-113">Os `name` e `description` aparecem no Excel e são localizados.</span><span class="sxs-lookup"><span data-stu-id="8b90c-113">The `name` and `description` appear in Excel and are localized.</span></span> <span data-ttu-id="8b90c-114">No entanto, o `id` de cada função não é localizado.</span><span class="sxs-lookup"><span data-stu-id="8b90c-114">However, the `id` of each function isn't localized.</span></span> <span data-ttu-id="8b90c-115">A `id` propriedade é como o Excel identifica sua função como exclusiva e não deve ser alterada depois de ser definida.</span><span class="sxs-lookup"><span data-stu-id="8b90c-115">The `id` property is how Excel identifies your function as unique and shouldn't be changed once it is set.</span></span>

<span data-ttu-id="8b90c-116">O JSON a seguir mostra como definir uma função com a `id` Propriedade "multiplique".</span><span class="sxs-lookup"><span data-stu-id="8b90c-116">The following JSON shows how to define a function with the `id` property "MULTIPLY."</span></span> <span data-ttu-id="8b90c-117">A `name` `description` propriedade e da função está localizada para alemão.</span><span class="sxs-lookup"><span data-stu-id="8b90c-117">The `name` and `description` property of the function is localized for German.</span></span> <span data-ttu-id="8b90c-118">Cada parâmetro `name` e `description` também é localizado para alemão.</span><span class="sxs-lookup"><span data-stu-id="8b90c-118">Each parameter `name` and `description` is also localized for German.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "SUMME",
    "description": "Summe zwei Zahlen",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "eins",
            "description": "Erste Nummer",
            "dimensionality": "scalar"
        },
        {
            "name": "zwei",
            "description": "Zweite Nummer",
            "dimensionality": "scalar"
        },
    ],
}
```

<span data-ttu-id="8b90c-119">Compare o JSON anterior com o seguinte JSON para inglês.</span><span class="sxs-lookup"><span data-stu-id="8b90c-119">Compare the previous JSON with the following JSON for English.</span></span>

```JSON
{
    "id": "MULTIPLY",
    "name": "Multiply",
    "description": "Multiplies two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "one",
            "description": "first number",
            "dimensionality": "scalar"
        },
        {
            "name": "two",
            "description": "second number",
            "dimensionality": "scalar"
        },
    ],
}
```

## <a name="localize-your-add-in"></a><span data-ttu-id="8b90c-120">Localizar seu suplemento</span><span class="sxs-lookup"><span data-stu-id="8b90c-120">Localize your add-in</span></span>

<span data-ttu-id="8b90c-121">Após criar um arquivo JSON para cada idioma, atualize o arquivo de manifesto XML com um valor de substituição para cada localidade que especifica a URL de cada arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="8b90c-121">After creating a JSON file for each language, update your XML manifest file with an override value for each locale that specifies the URL of each JSON metadata file.</span></span> <span data-ttu-id="8b90c-122">O seguinte XML de manifesto mostra uma `en-us` localidade padrão com uma URL de arquivo JSON de substituição para `de-de` (Alemanha).</span><span class="sxs-lookup"><span data-stu-id="8b90c-122">The following manifest XML shows a default `en-us` locale with an override JSON file URL for `de-de` (Germany).</span></span> <span data-ttu-id="8b90c-123">O arquivo de **funções-de. JSON** contém os nomes e IDs de função alemão localizados.</span><span class="sxs-lookup"><span data-stu-id="8b90c-123">The **functions-de.json** file contains the localized German function names and ids.</span></span>

```XML
<DefaultLocale>en-us</DefaultLocale>
...
<Resources>
     <bt:Urls>
        <bt:Url id="Contoso.Functions.Metadata.Url" DefaultValue="https://localhost:3000/dist/functions.json"/>
          <bt:Override Locale="de-de" Value="https://localhost:3000/dist/functions-de.json" />
        </bt:url>
        
     </bt:Urls>
</Resources>
```

<span data-ttu-id="8b90c-124">Para obter mais informações sobre o processo de localização de um suplemento, confira [localização para suplementos do Office](../develop/localization.md#control-localization-from-the-manifest).</span><span class="sxs-lookup"><span data-stu-id="8b90c-124">For more information on the process of localizing an add-in, see [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).</span></span>

## <a name="next-steps"></a><span data-ttu-id="8b90c-125">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="8b90c-125">Next steps</span></span>
<span data-ttu-id="8b90c-126">Saiba mais sobre [convenções de nomenclatura para funções personalizadas ou para](custom-functions-naming.md) descobrir [as práticas recomendadas de tratamento de erros](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="8b90c-126">Learn about [naming conventions for custom functions](custom-functions-naming.md) or discover [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8b90c-127">Confira também</span><span class="sxs-lookup"><span data-stu-id="8b90c-127">See also</span></span>

* [<span data-ttu-id="8b90c-128">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8b90c-128">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8b90c-129">Gerar metadados JSON automaticamente para funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8b90c-129">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="8b90c-130">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="8b90c-130">Create custom functions in Excel</span></span>](custom-functions-overview.md)
