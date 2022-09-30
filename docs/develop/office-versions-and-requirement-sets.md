---
title: Versões do Office e conjuntos de requisitos
description: Suporte a plataformas do Office.js usando API JavaScript.
ms.date: 09/14/2022
ms.localizationpriority: high
ms.openlocfilehash: 669977f87974a1ec5519ddbbe3d38c5a290ec84f
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234904"
---
# <a name="office-versions-and-requirement-sets"></a>Versões do Office e conjuntos de requisitos

Há várias versões do Office em várias plataformas, e nem todas dão suporte a cada API na API JavaScript para Office (Office.js). O Office 2013 no Windows era a versão mais antiga do Office que era compatível com suplementos do Office. Talvez você nem sempre tenha controle sobre a versão do Office que seus usuários instalaram. Para lidar com essa situação, fornecemos um sistema chamado conjuntos de requisitos para ajudá-lo a determinar se um aplicativo do Office dá suporte aos recursos necessários em seu Suplemento do Office.

> [!NOTE]
>
> - O Office pode ser executado em várias plataformas, incluindo o Windows, navegadores, Mac e iPad.
> - Exemplos de aplicativos do Office são: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.
> - O Office está disponível por uma assinatura do Microsoft 365 ou uma licença perpétua. A versão perpétua está disponível por contrato de licenciamento por volume ou varejo.
> - Um conjunto de requisitos é um grupo nomeado de membros da API, por exemplo, `ExcelApi 1.5`e `WordApi 1.3`assim por diante.

## <a name="how-to-check-your-office-version"></a>Como verificar sua versão do Office

Para identificar a versão do Office que você está usando, em um aplicativo do Office, selecione o menu **Arquivo** e escolha **Conta**. A versão do Office aparece na seção **Informações do** Produto. Por exemplo, a captura de tela a seguir indica o Office Versão 1802 (Build 9026.1000).

![Verificar sua versão do Office.](../images/office-version.png)

> [!NOTE]
> Se sua versão do Office for diferente disso, consulte Qual versão do Outlook eu tenho [?](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) ou Sobre o [Office:](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) Qual versão do Office estou usando? para entender como obter essas informações para sua versão.

## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Os Suplementos do Office podem usar conjuntos de requisitos de API para determinar se o aplicativo do Office dá suporte aos membros da API que ele precisa usar. O suporte ao conjunto de requisitos varia de acordo com o aplicativo do Office e a versão do aplicativo do Office (consulte a seção anterior [Como verificar sua versão do Office](#how-to-check-your-office-version)).

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

Além disso, outras funcionalidades como comandos de suplemento (extensibilidade da faixa de opções) e a capacidade de iniciar caixas de diálogo (API de Diálogo) foram adicionadas a API comum. Comandos de suplemento e conjuntos de requisitos da API de Caixa de Diálogo são exemplos de conjuntos de API que vários aplicativos do Office compartilham em comum.

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Conjuntos de requisitos de API JavaScript para Excel](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Conjuntos de requisitos da API JavaScript do Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (Caixa de Correio)
- [Conjuntos de requisitos da API JavaScript do PowerPoint](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Conjuntos de requisitos de API JavaScript para Word](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

Alguns conjuntos de requisitos contêm APIs que podem ser usadas por vários aplicativos do Office. Para obter informações sobre esses conjuntos de requisitos, consulte os artigos a seguir.

- [Conjuntos de requisitos comuns do Office](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [Conjuntos de requisitos dos comandos de suplemento](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [Conjuntos de requisitos da API de Caixa de Diálogo](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [Conjuntos de requisitos de origem da caixa de diálogo](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [Conjuntos de requisitos da API de Identidade](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [Conjuntos de requisitos de Coerção de Imagens](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [Conjuntos de requisitos de Atalhos de teclado](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [Abrir conjuntos de requisitos de janela do navegador](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [Conjuntos de requisitos comuns da API](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [Conjuntos de requisitos de tempo de execução compartilhado](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Especificar aplicativos do Office e conjuntos de requisitos

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>Confira também

- [Especificar requisitos da API e de aplicativos do Office](../develop/specify-office-hosts-and-api-requirements.md)
- [Instalar a última versão do Office](../develop/install-latest-office-version.md)
- [Visão geral dos canais de atualização do Microsoft 365 Apps](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Reinvente a produtividade com o Microsoft 365 e o Microsoft Teams](https://products.office.com/compare-all-microsoft-office-products?tab=2)
