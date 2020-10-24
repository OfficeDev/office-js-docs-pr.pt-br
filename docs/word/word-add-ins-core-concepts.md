---
title: Modelo de objeto de JavaScript do Word em Suplementos do Office
description: Aprenda as classes mais importantes no modelo de objeto de JavaScript específico do Word.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: c85c56987ef5de7c087064ac668f137326089642
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740865"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="19f29-103">Modelo de objeto de JavaScript do Word em Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="19f29-103">Word JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="19f29-104">Este artigo descreve conceitos fundamentais para o uso da [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) para criar suplementos. Ele introduz os principais conceitos fundamentais para o sua da API.</span><span class="sxs-lookup"><span data-stu-id="19f29-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins. It introduces core concepts that are fundamental to using the API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="19f29-105">Confira [Usar o modelo da API específica do aplicativo](../develop/application-specific-api-model.md) para saber mais sobre a natureza assíncrona das APIs do Word e como elas funcionam com o documento.</span><span class="sxs-lookup"><span data-stu-id="19f29-105">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.</span></span>

## <a name="officejs-apis-for-word"></a><span data-ttu-id="19f29-106">APIs Office.js para Word</span><span class="sxs-lookup"><span data-stu-id="19f29-106">Office.js APIs for Word</span></span>

<span data-ttu-id="19f29-107">Um suplemento do Word interage com objetos no Excel usando a API JavaScript do Office, que inclui dois modelos de objetos JavaScript:</span><span class="sxs-lookup"><span data-stu-id="19f29-107">A Word add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="19f29-108">**API JavaScript do Word**: a [API JavaScript do Word](../reference/overview/word-add-ins-reference-overview.md) fornece objetos fortemente tipados que você pode usar para acessar documentos, intervalos, tabelas, listas, formatação e mais.</span><span class="sxs-lookup"><span data-stu-id="19f29-108">**Word JavaScript API**: The [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access the document, ranges, tables, lists, formatting, and more.</span></span>

* <span data-ttu-id="19f29-109">**As APIs Comuns**: a [API Comum](/javascript/api/office) pode ser usada para acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="19f29-109">**Common APIs**: The [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="19f29-110">Embora você provavelmente usará a API JavaScript do Word para desenvolver a maioria das funcionalidades em suplementos que visam o Word, você também usará objetos na API comum.</span><span class="sxs-lookup"><span data-stu-id="19f29-110">While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API.</span></span> <span data-ttu-id="19f29-111">Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="19f29-111">For example:</span></span>

* <span data-ttu-id="19f29-112">[Context](/javascript/api/office/office.context): O `Context`objeto representa o ambiente de tempo de execução do suplemento e fornece acesso a objetos principais da API.</span><span class="sxs-lookup"><span data-stu-id="19f29-112">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="19f29-113">Ele consiste em configuração do documento, como `contentLanguage` e `officeTheme`, além de fornecer informações sobre o ambiente de tempo de execução do suplemento, como `host` e `platform`.</span><span class="sxs-lookup"><span data-stu-id="19f29-113">It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="19f29-114">Além disso, ele fornece o método `requirements.isSetSupported()`, que você pode usar para verificar se um conjunto de requisitos especificado é suportado pelo aplicativo Excel onde o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="19f29-114">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="19f29-115">[Documento](/javascript/api/office/office.document): o `Document` objeto fornece o método `getFileAsync()`, que você pode usar para baixar o arquivo do Word em que o suplemento está sendo executado.</span><span class="sxs-lookup"><span data-stu-id="19f29-115">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running.</span></span>

![Imagem das diferentes entre a API JS do Word e as APIs comuns](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a><span data-ttu-id="19f29-117">Modelo de objeto específico do Word</span><span class="sxs-lookup"><span data-stu-id="19f29-117">Word-specific object model</span></span>

<span data-ttu-id="19f29-118">Para entender as APIs do Word, você deve entender como os componentes de um documento estão relacionados entre si.</span><span class="sxs-lookup"><span data-stu-id="19f29-118">To understand the Word APIs, you must understand how the components of a document are related to one another.</span></span>

* <span data-ttu-id="19f29-119">O **documento** contém as **seções**, e entidades no nível de documento, como as configurações e partes XML Personalizadas.</span><span class="sxs-lookup"><span data-stu-id="19f29-119">The **Document** contains the **Section**s, and document-level entities such as settings and custom XML parts.</span></span>
* <span data-ttu-id="19f29-120">Uma **seção** contém um**corpo**.</span><span class="sxs-lookup"><span data-stu-id="19f29-120">A **Section** contains a **Body**.</span></span>
* <span data-ttu-id="19f29-121">Um **corpo** dá acesso a **parágrafo**s, **ContentControl**s e aos objetos do **intervalo**, entre outros.</span><span class="sxs-lookup"><span data-stu-id="19f29-121">A **Body** gives access to **Paragraph**s, **ContentControl**s, and **Range** objects, among others.</span></span>
* <span data-ttu-id="19f29-122">Um **intervalo** representa uma área contínua de conteúdo, incluindo texto, espaço em branco, **tabela**s e imagens.</span><span class="sxs-lookup"><span data-stu-id="19f29-122">A **Range** represents a contiguous area of content, including text, white space, **Table**s, and images.</span></span> <span data-ttu-id="19f29-123">Ele também contém a maioria dos métodos de manipulação de texto.</span><span class="sxs-lookup"><span data-stu-id="19f29-123">It also contains most of the text manipulation methods.</span></span>
* <span data-ttu-id="19f29-124">Uma **Lista** representa o texto em uma lista numerada ou em lista com marcadores.</span><span class="sxs-lookup"><span data-stu-id="19f29-124">A **List** represents text in a numbered or bulleted list.</span></span>

## <a name="see-also"></a><span data-ttu-id="19f29-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="19f29-125">See also</span></span>

- [<span data-ttu-id="19f29-126">Visão geral da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="19f29-126">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="19f29-127">Criar seu primeiro suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="19f29-127">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="19f29-128">Tutorial de suplemento do Word</span><span class="sxs-lookup"><span data-stu-id="19f29-128">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="19f29-129">Referências da API JavaScript do Word</span><span class="sxs-lookup"><span data-stu-id="19f29-129">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="19f29-130">Saiba mais sobre o Programa de Desenvolvedores do Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="19f29-130">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)