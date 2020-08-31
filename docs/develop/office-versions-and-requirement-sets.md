---
title: Versões do Office e conjuntos de requisitos
description: Suporte a plataformas do Office.js usando API JavaScript
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 14b88402b1ee563d992b6f37f95be4fa7f337388
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293209"
---
# <a name="office-versions-and-requirement-sets"></a>Versões do Office e conjuntos de requisitos

Há várias versões do Office em várias plataformas, e nem todas dão suporte a cada API na API JavaScript para Office (Office.js). Nem sempre você terá controle sobre a versão do Office que os usuários instalaram.  Para lidar com essa situação, fornecemos um sistema chamado conjuntos de requisitos para ajudar você a determinar se um aplicativo do Office dá suporte aos recursos necessários em seu Suplemento do Office. 

> [!NOTE]
> - O Office pode ser executado em várias plataformas, incluindo o Windows, navegadores, Mac e iPad.
> - Entre os exemplos dos aplicativos do Office estão os produtos do Office: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.  
> - Um conjunto de requisito é um grupo nomeado de membros da API, por exemplo, `ExcelApi 1.5`, `WordApi 1.3` etc.  

## <a name="how-to-check-your-office-version"></a>Como verificar sua versão do Office

Para identificar a versão do Office que você está usando, em um aplicativo do Office, selecione o menu **Arquivo** e escolha **Conta**. A versão do Office aparecerá na seção **Informações do Produto**. Por exemplo, a captura de tela a seguir indica o Office Versão 1802 (Build 9026.1000):

![Verificar sua versão do Office](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Os Suplementos do Office podem usar conjuntos de requisitos de API para determinar se o aplicativo do Office oferece suporte aos membros da API necessários. O suporte a um conjunto de requisitos varia de acordo com o aplicativo do Office e a versão do aplicativo do Office (veja a seção anterior).

Alguns aplicativos do Office possuem seus próprios conjuntos de requisitos de API. Por exemplo, o primeiro conjunto de requisitos para a API do Excel foi `ExcelApi 1.1`, e o primeiro conjunto de requisitos para a API do Word foi `WordApi 1.1`. Desde então, vários conjuntos de requisitos novos de ExcelApi e WordApi foram adicionados para fornecer mais funcionalidades de API.

Além disso, outras funcionalidades como comandos de suplemento (extensibilidade da faixa de opções) e a capacidade de iniciar caixas de diálogo (API de Diálogo) foram adicionadas a API comum. Os comandos de suplemento e os conjuntos de requisitos de API de Diálogo são exemplos de conjuntos de API que os diversos aplicativos do Office compartilham em comum.

Um suplemento só pode usar APIs em conjuntos de requisitos compatíveis com a versão do aplicativo do Office na qual ele está em execução. Para saber exatamente quais conjuntos de requisitos estão disponíveis para uma versão específica do aplicativo do Office, confira os seguintes artigos sobre conjunto de requisitos específicos ao aplicativo:

- [Conjuntos de requisitos de API JavaScript para Excel](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [Conjuntos de requisitos da API JavaScript do PowerPoint](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [Noções básicas sobre os conjuntos de requisitos da API do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md) (Caixa de Correio)

Alguns conjuntos de requisito contêm APIs que podem ser usadas por qualquer aplicativo do Office. Para saber mais sobre esses conjuntos de requisitos, confira estes artigos:

- [Conjuntos de requisitos comuns do Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Conjuntos de requisitos dos comandos de suplemento](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Conjuntos de requisitos da API de Caixa de Diálogo](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Conjuntos de requisitos da API de Identidade](../reference/requirement-sets/identity-api-requirement-sets.md)

O número da versão de um conjunto de requisitos, como "1.1" no `ExcelApi 1.1`, tem relação com o aplicativo do Office. O número da versão de um certo conjunto de requisitos (por exemplo, `ExcelApi 1.1`), não corresponde ao número da versão do Office.js ou aos conjuntos de requisitos para outros aplicativos do Office (por exemplo, Word, Outlook etc.).  Lançamos os conjuntos de requisitos para diferentes aplicativos do Office em ritmos e períodos diferentes. Por exemplo, `ExcelApi 1.5` foi lançado antes do conjunto de requisitos `WordApi 1.3`.


A biblioteca da API JavaScript do Office (Office.js) inclui todos os conjuntos de requisitos disponíveis no momento. Embora exista algo como conjuntos de requisitos `ExcelApi 1.3` e `WordApi 1.3`, há nenhum conjunto de requisitos `Office.js 1.3`. A versão mais recente do Office.js é mantida como um único ponto de extremidade do Office fornecida por meio da CDN (rede de distribuição de conteúdo). Para obter mais detalhes sobre a CDN do Office.js, incluindo como a versão e a compatibilidade com versões anteriores são tratadas, consulte [Noções básicas sobre a API JavaScript do Office](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>Especificar aplicativos do Office e conjuntos de requisitos

Há várias maneiras de especificar quais aplicativos do Office e conjuntos de requisitos são exigidos por um suplemento.  Para saber mais detalhes, confira [Especificar requisitos de API e aplicativos do Office](../develop/specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>Confira também

- [Especificar requisitos da API e de aplicativos do Office](../develop/specify-office-hosts-and-api-requirements.md)
- [Instalar a última versão do Office](../develop/install-latest-office-version.md)
- [Visão geral dos canais de atualização do Microsoft 365 Apps](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirar o máximo proveito do Office com o Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
