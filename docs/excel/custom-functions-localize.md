---
ms.date: 11/06/2020
description: Localize suas Excel funções personalizadas.
title: Localize funções personalizadas
ms.localizationpriority: medium
ms.openlocfilehash: 596ab23f578f7bb0d12d009d06871e946302300c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148938"
---
# <a name="localize-custom-functions"></a>Localize funções personalizadas

Você pode localizar seu complemento e seus nomes de função personalizados. Para fazer isso, forneça nomes de função localizados no arquivo JSON das funções e informações de localidade no arquivo de manifesto XML.

>[!IMPORTANT]
> Os metadados gerados automaticamente não funcionam para localização, portanto, você precisa atualizar o arquivo JSON manualmente. Para saber como fazer isso, consulte [Manualmente criar metadados JSON para funções personalizadas](custom-functions-json.md)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="localize-function-names"></a>Nomes de função de localização

Para localizar suas funções personalizadas, crie um novo arquivo de metadados JSON para cada idioma. Em cada arquivo JSON de idioma, crie `name` e propriedades no idioma de `description` destino. O arquivo padrão para inglês é **chamado functions.json**. Use a localidade no nome do arquivo para cada arquivo JSON adicional, como **functions-de.json** para ajudar a identificá-los.

The `name` and appear in Excel and are `description` localized. No entanto, `id` a de cada função não está localizada. A propriedade é como Excel identifica sua função como exclusiva e não deve `id` ser alterada depois de definida.

O JSON a seguir mostra como definir uma função com a `id` propriedade "MULTIPLY". A `name` propriedade e da função é localizada para `description` alemão. Cada parâmetro `name` e também é localizado para `description` alemão.

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

Compare o JSON anterior com o JSON a seguir para inglês.

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

## <a name="localize-your-add-in"></a>Localize seu add-in

Depois de criar um arquivo JSON para cada idioma, atualize seu arquivo de manifesto XML com um valor de substituição para cada localidade que especifica a URL de cada arquivo de metadados JSON. O XML de manifesto a seguir mostra uma `en-us` localidade padrão com uma URL de arquivo JSON de substituição para `de-de` (Alemanha). O **arquivo functions-de.json** contém os nomes e as ids de funções alemãs localizadas.

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

Para obter mais informações sobre o processo de localização de um add-in, consulte [Localization for Office Add-ins](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Próximas etapas
Saiba mais [sobre as convenções de nomenis para funções personalizadas](custom-functions-naming.md) ou descubra as [práticas recomendadas de](custom-functions-errors.md)tratamento de erros.

## <a name="see-also"></a>Confira também

* [Criar metadados JSON manualmente para funções personalizadas](custom-functions-json.md)
* [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
