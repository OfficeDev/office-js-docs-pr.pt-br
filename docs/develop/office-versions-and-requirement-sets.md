---
title: Versões do Office e conjuntos de requisitos
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505990"
---
# <a name="office-versions-and-requirement-sets"></a>Versões do Office e conjuntos de requisitos

Existem várias versões do Office em várias plataformas, mas nem todas oferecem suporte para todas as APIs JavaScript do Office (Office.js). Pode ser que nem sempre você tenha controle sobre a versão instalada pelos usuários. Para lidar com isso, nós fornecemos um sistema chamado "conjuntos de requisitos" para ajudar você a determinar se um host do Office suporta os recursos necessários para o seu suplemento do Office. 

> [!NOTE]
> - O Office é executado em várias plataformas, incluindo o Office Online, para Windows, Mac e iPad.  
> - Alguns exemplos de hosts do Office são os próprios produtos do Office: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.  
> - Um conjunto de requisitos é um grupo nomeado de membros de uma API, por exemplo, `ExcelApi 1.5`, `WordApi 1.3`, etc.  


## <a name="how-to-check-your-office-version"></a>Como verificar sua versão do Office

Para identificar a versão do Office que você está usando, de dentro de um aplicativo do Office, selecione o menu **Arquivo** e acesse **Conta**. A versão do Office aparecerá na seção **Informações do produto**. Por exemplo, a captura de tela a seguir indica de versão 1802 do Office (Build 9026.1000):

![Verificar sua versão do Office](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Os suplementos do Office podem usar conjuntos de requisitos de API para determinar se o host do Office suporta os membros necessários da API. O suporte dos conjuntos de requisitos varia de acordo com o host do Office e a versão dele (consulte a seção anterior).

Alguns hosts do Office têm seus próprios conjuntos de requisito de API. Por exemplo, o primeiro conjunto de requisitos definido para a API do Excel foi `ExcelApi 1.1`, já para a API do Word foi `WordApi 1.1`. Desde então, vários novos conjuntos de requisitos foram adicionados à API do Excele e do Word para fornecer funcionalidades adicionais.

Além disso, outras funcionalidades, como comandos de suplemento (extensibilidade da faixa de opções) e a capacidade de iniciar caixas de diálogo (API de diálogo), foram adicionadas à API comum. Os conjuntos de requisitos de comandos de suplemento e de APIs de diálogo são exemplos de conjuntos de APIs que os vários hosts do Office compartilham em comum.

Um suplemento só pode usar APIs em conjuntos de requisitos suportados pela versão do host do Office em que o suplemento é executado. Para saber exatamente quais conjuntos de requisitos estão disponíveis para uma versão específica de host do Office, consulte os seguintes artigos:

- [Conjuntos de requisitos da API JavaScript do Excel](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)
- [Conjuntos de requisitos da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)
- [Noções básicas sobre conjuntos de requisitos da API do Outlook](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)

Alguns conjuntos de requisitos contêm APIs que podem ser usadas por qualquer host do Office. Para obter informações sobre esses conjuntos, consulte os seguintes artigos:

- [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [Conjuntos de requisitos de comandos de suplemento](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [Conjuntos de requisitos da API de caixa de diálogo](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Identificar conjuntos de requisitos de API](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

O número de versão de um conjunto de requisitos, como "1.1" para `ExcelApi 1.1`, é relativo ao host do Office. O número da versão de um conjunto de requisitos específico (por exemplo, `ExcelApi 1.1`) não corresponde ao número da versão do Office.js, nem ao de conjuntos de requisitos de outros hosts do Office (por exemplo, Word, Outlook, etc.).  Conjuntos de requisitos de diferentes hosts do Office são lançados em diferentes velocidades e horários. Por exemplo, `ExcelApi 1.5` foi lançado antes do conjunto de requisitos `WordApi 1.3`.

A  biblioteca da API JavaScript para Office (Office. js) inclui todos os conjuntos de requisitos atualmente disponíveis. Apesar de haver algo como conjuntos de requisitos `ExcelApi 1.3` e `WordApi 1.3`, não há nenhum conjunto de requisitos `Office.js 1.3`. A versão mais recente do Office.js é mantida como um único ponto de extremidade do Office fornecido por meio da rede de fornecimento de conteúdo (CDN). Para obter mais detalhes sobre a CDN do Office.js, incluindo como funcionam o controle de versão e compatibilidade com versões anteriores, consulte [Noções básicas sobre a API JavaScript para Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

## <a name="specify-office-hosts-and-requirement-sets"></a>Especificar conjuntos de requisitos e hosts do Office

Há várias maneiras de especificar quais conjuntos de requisitos e hosts do Office são exigidos por um suplemento.  Para obter informações detalhadas, consulte [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)


## <a name="see-also"></a>Confira também

- [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Instalar a última versão do Office](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [Visão geral dos canais de atualização do Office 365 ProPlus](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Aproveitar ao máximo o Office com o Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
