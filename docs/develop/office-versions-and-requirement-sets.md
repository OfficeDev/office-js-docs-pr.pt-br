---
title: Versões do Office e conjuntos de requisitos
description: Suporte a plataformas do Office.js usando API JavaScript
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 02f3d91256ea05e526ebe2e4e4090b1908d7292a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093578"
---
# <a name="office-versions-and-requirement-sets"></a>Versões do Office e conjuntos de requisitos

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in. 

> [!NOTE]
> - O Office pode ser executado em várias plataformas, incluindo o Windows, navegadores, Mac e iPad.
> - Entre os exemplos de hosts do Office estão os produtos do Office: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.  
> - Um conjunto de requisito é um grupo nomeado de membros da API, por exemplo, `ExcelApi 1.5`, `WordApi 1.3` etc.  

## <a name="how-to-check-your-office-version"></a>Como verificar sua versão do Office

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):

![Verificar sua versão do Office](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).

Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Além disso, outras funcionalidades como comandos de suplemento (extensibilidade da faixa de opções) e a capacidade de iniciar caixas de diálogo (API de Diálogo) foram adicionadas a API comum. Os comandos de suplemento e os conjuntos de requisitos de API de Diálogo são exemplos de conjuntos de API que os diversos hosts do Office compartilham em comum.

An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:

- [Conjuntos de requisitos de API JavaScript para Excel](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [Conjuntos de requisitos da API JavaScript do PowerPoint](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Noções básicas sobre os conjuntos de requisitos da API do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md) (Caixa de Correio)

Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:

- [Conjuntos de requisitos comuns do Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Conjuntos de requisitos dos comandos de suplemento](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Conjuntos de requisitos da API de Identidade](../reference/requirement-sets/identity-api-requirement-sets.md)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

A biblioteca da API JavaScript do Office (Office.js) inclui todos os conjuntos de requisitos disponíveis no momento. Embora exista algo como conjuntos de requisitos `ExcelApi 1.3` e `WordApi 1.3`, há nenhum conjunto de requisitos `Office.js 1.3`. A versão mais recente do Office.js é mantida como um único ponto de extremidade do Office fornecida por meio da CDN (rede de distribuição de conteúdo). Para obter mais detalhes sobre a CDN do Office.js, incluindo como a versão e a compatibilidade com versões anteriores são tratadas, consulte [Noções básicas sobre a API JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-hosts-and-requirement-sets"></a>Especificar hosts do Office e conjuntos de requisitos

There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>Confira também

- [Especificar requisitos da API e de hosts do Office](../develop/specify-office-hosts-and-api-requirements.md)
- [Instalar a última versão do Office](../develop/install-latest-office-version.md)
- [Visão geral dos canais de atualização do Microsoft 365 Apps](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirar o máximo proveito do Office com o Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
