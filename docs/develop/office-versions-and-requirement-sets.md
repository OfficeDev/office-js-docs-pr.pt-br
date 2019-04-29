---
title: Versões do Office e conjuntos de requisitos
description: ''
ms.date: 04/19/2019
localization_priority: Priority
ms.openlocfilehash: e1047501cdac8dc88ab9f7778b846e171ee02d44
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/28/2019
ms.locfileid: "33440035"
---
# <a name="office-versions-and-requirement-sets"></a>Versões do Office e conjuntos de requisitos

Há várias versões do Office em várias plataformas, e nem todas dão suporte a cada API na API JavaScript para Office (Office.js). Nem sempre você terá controle sobre a versão do Office que os usuários instalaram.  Para lidar com essa situação, fornecemos um sistema chamado de conjuntos de requisitos para ajudar você a determinar se um host do Office dá suporte para com os recursos necessários em seu Suplemento do Office. 

> [!NOTE]
> - O Office é executado em várias plataformas, incluindo o Office para Windows, o Office Online, o Office para Mac e o Office para iPad.
> - Entre os exemplos de hosts do Office estão os produtos do Office: Excel, Word, PowerPoint, Outlook, OneNote e assim por diante.  
> - Um conjunto de requisito é um grupo nomeado de membros da API, por exemplo, `ExcelApi 1.5`, `WordApi 1.3` etc.  


## <a name="how-to-check-your-office-version"></a>Como verificar sua versão do Office

Para identificar a versão do Office que você está usando, em um aplicativo do Office, selecione o menu **Arquivo** e escolha **Conta**. A versão do Office aparecerá na seção **Informações do Produto**. Por exemplo, a captura de tela a seguir indica o Office Versão 1802 (Build 9026.1000):

![Verificar sua versão do Office](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Disponibilidade dos conjuntos de requisitos do Office

Os Suplementos do Office podem usar conjuntos de requisitos de API para determinar se o host do Office oferece suporte aos membros da API necessários. O suporte a um conjunto de requisitos varia de acordo com o host do Office e a versão do host do Office (veja a seção anterior).

Alguns hosts do Office tem seus próprios conjuntos de requisitos de API. Por exemplo, o primeiro conjunto de requisitos para a API do Excel foi `ExcelApi 1.1`, e o primeiro conjunto de requisitos para a API do Word foi `WordApi 1.1`. Desde então, vários conjuntos de requisitos novos de ExcelApi e WordApi foram adicionados para fornecer mais funcionalidades de API.

Além disso, outras funcionalidades como comandos de suplemento (extensibilidade da faixa de opções) e a capacidade de iniciar caixas de diálogo (API de Diálogo) foram adicionadas a API comum. Os comandos de suplemento e os conjuntos de requisitos de API de Diálogo são exemplos de conjuntos de API que os diversos hosts do Office compartilham em comum.

Um suplemento só pode usar APIs em conjuntos de requisitos compatíveis com a versão do host do Office na qual ele está em execução. Para saber exatamente quais conjuntos de requisitos estão disponíveis para uma versão específica de host do Office, confira os seguintes artigos sobre conjunto de requisitos específicos ao host:

- [Conjuntos de requisitos de API JavaScript para Excel](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) (ExcelApi)
- [Conjuntos de requisitos de API JavaScript para Word](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)
- [Conjuntos de requisitos de API JavaScript para OneNote](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)
- [Noções básicas sobre os conjuntos de requisitos da API do Outlook](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (Caixa de Correio)

Alguns conjuntos de requisito contêm APIs que podem ser usadas por qualquer host do Office. Para saber mais sobre esses conjuntos de requisitos, confira estes artigos:

- [Conjuntos de requisitos comuns do Office](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [Conjuntos de requisitos dos comandos de suplemento](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [Conjuntos de requisitos da API de Caixa de Diálogo](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [Conjuntos de requisitos da API de Identidade](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)

O número da versão de um conjunto de requisitos, como "1.1" no `ExcelApi 1.1`, tem relação com o host do Office. O número da versão de um certo conjunto de requisitos (por exemplo, `ExcelApi 1.1`), não corresponde ao número da versão do Office.js ou aos conjuntos de requisitos para outros hosts do Office (por exemplo, Word, Outlook etc.).  Lançamos os conjuntos de requisitos para diferentes hosts do Office em ritmos e períodos diferentes. Por exemplo, `ExcelApi 1.5` foi lançado antes do conjunto de requisitos `WordApi 1.3`.

A biblioteca da API JavaScript para Office (Office.js) inclui todos os conjuntos de requisitos disponíveis no momento. Embora exista algo como conjuntos de requisitos `ExcelApi 1.3` e `WordApi 1.3`, há nenhum conjunto de requisitos `Office.js 1.3`. A versão mais recente do Office.js é mantida como um único ponto de extremidade do Office fornecida por meio da CDN (rede de distribuição de conteúdo). Saiba mais sobre a CDN do Office.js, inclusive como é feito o controle de versão e como lidar com a compatibilidade com versões anteriores, em [Noções básicas da API JavaScript para Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).

## <a name="specify-office-hosts-and-requirement-sets"></a>Especificar hosts do Office e conjuntos de requisitos

Há várias maneiras de especificar quais hosts do Office e conjuntos de requisitos são exigidos por um suplemento.  Para saber mais detalhes, confira [Especificar requisitos de API e hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements).


## <a name="see-also"></a>Confira também

- [Especificar requisitos da API e de hosts do Office](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Instalar a última versão do Office](/office/dev/add-ins/develop/install-latest-office-version)
- [Visão geral dos canais de atualização do Office 365 ProPlus](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Tirar o máximo proveito do Office com o Office 365](https://products.office.com/compare-all-microsoft-office-products?tab=2)
