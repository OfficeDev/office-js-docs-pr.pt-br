---
ms.date: 06/17/2019
description: Localize suas funções personalizadas do Excel.
title: Localizar funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 7c289f65a7d75f1c1c07770d43e09f92568ca73b
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059710"
---
# <a name="localize-custom-functions"></a>Localizar funções personalizadas

Você pode localizar o suplemento e seus nomes de funções personalizadas. Você precisa fornecer nomes de função localizados no arquivo JSON funções e fornecer informações de localidade no arquivo de manifesto XML.

>[!IMPORTANT]
> Os metadados gerados automaticamente não funcionam para localização, portanto, você precisa atualizar o arquivo JSON manualmente.

## <a name="localize-function-names"></a>Localizar nomes de função

Para localizar suas funções personalizadas, crie um novo arquivo de metadados JSON para cada idioma. Em cada arquivo JSON de idioma, `name` crie `description` e propriedades no idioma de destino. O arquivo padrão para inglês é chamado de **funções. JSON**. É recomendável usar a localidade no nome do arquivo para cada arquivo JSON adicional, como **funções-de. JSON** para ajudá-lo a identificá-los.

Os `name` e `description` aparecem no Excel e são localizados. No entanto `id` , o de cada função não é localizado. A `id` propriedade é como o Excel identifica sua função como exclusiva e não deve ser alterada depois de ser definida.

O JSON a seguir mostra como definir uma função com a `id` Propriedade "multiplique". A `name` propriedade `description` e da função está localizada para alemão. Cada parâmetro `name` e `description` também é localizado para alemão.

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

Compare o JSON anterior com o seguinte JSON para inglês.

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

## <a name="localize-your-add-in"></a>Localizar seu suplemento

Após criar um arquivo JSON para cada idioma, você precisa atualizar seu arquivo de manifesto XML com um valor de substituição para cada localidade que especifica a URL de cada arquivo de metadados JSON. O seguinte XML de manifesto mostra uma `en-us` localidade padrão com uma URL de arquivo JSON `de-de` de substituição para (Alemanha). O arquivo de **funções-de. JSON** contém os nomes e IDs de função alemão localizados.

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

Para obter mais informações sobre o processo de localização de um suplemento, confira [localização para suplementos do Office](../develop/localization.md#control-localization-from-the-manifest).

## <a name="next-steps"></a>Próximas etapas
Saiba mais sobre [convenções de nomenclatura para funções personalizadas ou para](custom-functions-naming.md) descobrir [as práticas recomendadas de tratamento de erros](custom-functions-errors.md).

## <a name="see-also"></a>Confira também

* [Metadados de funções personalizadas](custom-functions-json.md)
* [Gerar metadados JSON automaticamente para funções personalizadas](custom-functions-json-autogeneration.md)
* [Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).
* [Criar funções personalizadas no Excel](custom-functions-overview.md)
