---
title: Vers?es do Office e conjuntos de requisitos
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: fe02a63e93bd7fbb8a2709b1e3977fee999e5b9a
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-versions-and-requirement-sets"></a>Vers?es do Office e conjuntos de requisitos

H? v?rias vers?es do Office em v?rias plataformas, e nem todas d?o suporte a cada API na API JavaScript para Office (Office.js). Nem sempre voc? ter? controle sobre a vers?o do Office que os usu?rios instalaram.  Para lidar com essa situa??o, fornecemos um sistema chamado de conjuntos de requisitos para ajudar voc? a determinar se um host do Office d? suporte para com os recursos necess?rios em seu Suplemento do Office. 

> [!NOTE]
> - O Office ? executado em v?rias plataformas, incluindo o Office para Windows, o Office Online, o Office para Mac e o Office para iPad.  
> - Entre os exemplos de hosts do Office est?o os produtos do Office: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.  
> - Um conjunto de requisito ? um grupo nomeado de membros da API, por exemplo, `ExcelApi 1.5`, `WordApi 1.3` etc.  


## <a name="how-to-check-your-office-version"></a>Como verificar sua vers?o do Office

Para identificar a vers?o do Office que voc? est? usando, em um aplicativo do Office, selecione o menu **Arquivo** e escolha **Conta**. A vers?o do Office aparecer? na se??o **Informa??es do Produto**. Por exemplo, a captura de tela a seguir indica o Office Vers?o 1802 (Build 9026.1000):

![Verificar sua vers?o do Office](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Os Suplementos do Office podem usar conjuntos de requisitos de API para determinar se o host do Office oferece suporte aos membros da API necess?rios. O suporte a um conjunto de requisitos varia de acordo com o host do Office e a vers?o do host do Office (veja a se??o anterior).

Alguns hosts do Office tem seus pr?prios conjuntos de requisitos de API. Por exemplo, o primeiro conjunto de requisitos para a API do Excel foi `ExcelApi 1.1`, e o primeiro conjunto de requisitos para a API do Word foi `WordApi 1.1`. Desde ent?o, v?rios conjuntos de requisitos novos de ExcelApi e WordApi foram adicionados para fornecer mais funcionalidades de API.

Al?m disso, outras funcionalidades como comandos de suplemento (extensibilidade da faixa de op??es) e a capacidade de iniciar caixas de di?logo (API de Di?logo) foram adicionadas a API comum. Os comandos de suplemento e os conjuntos de requisitos de API de Di?logo s?o exemplos de conjuntos de API que os diversos hosts do Office compartilham em comum.

Um suplemento s? pode usar APIs em conjuntos de requisitos compat?veis com a vers?o do host do Office na qual ele est? em execu??o. Para saber exatamente quais conjuntos de requisitos est?o dispon?veis para uma vers?o espec?fica de host do Office, confira os seguintes artigos sobre conjunto de requisitos espec?ficos ao host:

- [Conjuntos de requisitos de API JavaScript para Excel](https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets?product=excel) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets) (WordApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)
- [No??es b?sicas sobre conjuntos de requisitos da API do Outlook](https://dev.office.com/reference/add-ins/outlook/tutorial-api-requirement-sets) (MailBox)

Alguns conjuntos de requisito cont?m APIs que podem ser usadas por qualquer host do Office. Para saber mais sobre esses conjuntos de requisitos, confira estes artigos:

- [Conjuntos de requisitos comuns do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Conjuntos de requisitos dos comandos de suplemento](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets?product=excel)
- [Conjuntos de requisitos da API de Caixa de Di?logo](https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets?product=excel)
- [Conjuntos de requisitos da API de Identidade](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets?product=excel)

O n?mero da vers?o de um conjunto de requisitos, como "1.1" no `ExcelApi 1.1`, tem rela??o com o host do Office. O n?mero da vers?o de um certo conjunto de requisitos (por exemplo, `ExcelApi 1.1`), n?o corresponde ao n?mero da vers?o do Office.js ou aos conjuntos de requisitos para outros hosts do Office (por exemplo, Word, Outlook etc.).  Lan?amos os conjuntos de requisitos para diferentes hosts do Office em ritmos e per?odos diferentes. Por exemplo, `ExcelApi 1.5` foi lan?ado antes do conjunto de requisitos `WordApi 1.3`.

A biblioteca da API JavaScript para Office (Office.js) inclui todos os conjuntos de requisitos dispon?veis no momento. Embora exista algo como conjuntos de requisitos `ExcelApi 1.3` e `WordApi 1.3`, h? nenhum conjunto de requisitos `Office.js 1.3`. A vers?o mais recente do Office.js ? mantida como um ?nico ponto de extremidade do Office fornecido por meio da rede de distribui??o de conte?do (CDN). Saiba mais sobre a CDN do Office.js, inclusive como ? feito o controle de vers?o e como lidar com a compatibilidade com vers?es anteriores, em [No??es b?sicas da API JavaScript para Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

## <a name="specify-office-hosts-and-requirement-sets"></a>Especificar hosts do Office e conjuntos de requisitos

H? v?rias maneiras de especificar quais hosts do Office e conjuntos de requisitos s?o exigidos por um suplemento.  Para saber mais detalhes, confira [Especificar requisitos de API e hosts do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).


## <a name="see-also"></a>Veja tamb?m

- [Especificar requisitos da API e de hosts do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Instalar a ?ltima vers?o do Office](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/install-latest-office-version)
- [Vis?o geral dos canais de atualiza??o do Office 365 ProPlus](https://docs.microsoft.com/en-us/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirar o m?ximo proveito do Office com o Office 365](https://products.office.com/en-us/compare-all-microsoft-office-products?tab=2)
