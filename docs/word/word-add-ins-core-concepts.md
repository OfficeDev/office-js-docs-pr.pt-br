---
title: Modelo de objeto de JavaScript do Word em Suplementos do Office
description: Saiba mais sobre os principais componentes no modelo de objeto JavaScript específico do Word.
ms.date: 3/17/2022
ms.localizationpriority: high
ms.openlocfilehash: 07055ee2c8b16315b5c4efea5f62a85331e48445
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958955"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a>Modelo de objeto de JavaScript do Word em Suplementos do Office

Este artigo descreve os conceitos fundamentais para usar a [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos.

> [!IMPORTANT]
> Confira [Usar o modelo da API específica do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre a natureza assíncrona das APIs do Word e como elas funcionam com o documento.

## <a name="officejs-apis-for-word"></a>APIs Office.js para Word

Um suplemento do Word interage com objetos no Word usando a API JavaScript do Office. Isso inclui dois modelos de objeto JavaScript:

* **API JavaScript do Word**: a [API JavaScript do Word](/javascript/api/word) fornece objetos fortemente tipados que você pode usar para acessar documentos, intervalos, tabelas, listas, formatação e mais.

* **APIs comuns**: a [API Comum](/javascript/api/office) dá acesso a recursos como interface do usuário, caixas de diálogo e configurações de cliente que são comuns em vários aplicativos do Office.

Embora você provavelmente use a API JavaScript do Word para desenvolver a maioria das funcionalidades em suplementos destinados ao Word, você também usará objetos na API Comum. Por exemplo:

* [Office.Context](/javascript/api/office/office.context): o objeto `Context` representa o ambiente de tempo de execução do suplemento e fornece acesso aos principais objetos da API. Ele consiste em detalhes de configuração do documento, como `contentLanguage` e `officeTheme`, e também fornece informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`. Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se um conjunto de requisitos especificado é compatível com o aplicativo Word em que o suplemento está sendo executado.
* [Office.Document](/javascript/api/office/office.document): o objeto `Office.Document` fornece o método `getFileAsync()`, que você pode usar para fazer download do arquivo do Word onde o suplemento está sendo executado. Isso é separado do objeto [Word.Document](/javascript/api/word/word.document).

![Diferenças entre a API JS do Word e as APIs comuns.](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a>Modelo de objeto específico do Word

Para entender as APIs do Word, você deve entender como os componentes de um documento estão relacionados entre si.

* O **documento** contém as **seções**, e entidades no nível de documento, como as configurações e partes XML Personalizadas.
* Uma **seção** contém um **corpo**.
* Um **corpo** dá acesso a **parágrafo** s, **ContentControl** s e aos objetos do **intervalo**, entre outros.
* Um **intervalo** representa uma área contínua de conteúdo, incluindo texto, espaço em branco, **tabela** s e imagens. Ele também contém a maioria dos métodos de manipulação de texto.
* Uma **Lista** representa o texto em uma lista numerada ou em lista com marcadores.

## <a name="see-also"></a>Confira também

* [Visão geral da API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md)
* [Criar seu primeiro suplemento do Word](../quickstarts/word-quickstart.md)
* [Tutorial de suplemento do Word](../tutorials/word-tutorial.md)
* [Referências da API JavaScript do Word](/javascript/api/word)
* [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
