---
title: Fabric Core em Office de complementos
description: Obter uma visão geral de como usar o Fabric Core e os componentes da interface do usuário do Fabric em Office de complementos.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: e93efaea55841cc3bb6fa79ea1d1bbcaa76a4d05
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330196"
---
# <a name="fabric-core-in-office-add-ins"></a>Fabric Core em Office de complementos

Fabric Core é uma coleção open-source de classes CSS e mixins SASS que se destinam a ser usadas em React *Office* Add-ins. O Fabric Core contém elementos básicos da linguagem de design da interface do usuário fluente, como ícones, cores, tipos e grades. O Fabric Core é independente da estrutura, portanto, pode ser usado com qualquer aplicativo de página única ou qualquer estrutura de interface do usuário web do lado do servidor. (Chama-se "Fabric Core" em vez de "Fluent Core" por motivos históricos.)

Se a interface do usuário do seu React não for baseada em React, você também poderá usar um conjunto de componentes que não React. Consulte [Usar Office UI Fabric componentes JS](#use-office-ui-fabric-js-components).

> [!NOTE]
> Este artigo descreve o uso do Fabric Core no contexto de Office de complementos. Mas também é usado em uma ampla variedade de Microsoft 365 aplicativos e extensões. Para obter mais informações, [consulte Fabric Core](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core) and the open source repo Office UI Fabric [Core](https://github.com/OfficeDev/office-ui-fabric-core).

## <a name="use-fabric-core-icons-fonts-colors"></a>Uso do Fabric Core: ícones, fontes, cores

Para começar a usar o Fabric Core:

1. Adicione a referência da CDN ao HTML da sua página.  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Use ícones e fontes do Fabric Core.

    Para usar um ícone do Fabric Core, inclua o elemento "i" em sua página e, em seguida, fazer referência às classes apropriadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte. Por exemplo, o código a seguir mostra como criar um ícone de tabela muito grande que usa a cor themePrimary (#0078d7).

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    Para obter instruções mais detalhadas, consulte [Fluent UI Icons](https://developer.microsoft.com/fluentui#/styles/web/icons). Para encontrar mais ícones disponíveis no Fabric Core, use o recurso de pesquisa nessa página. Quando encontrar um ícone para usar no suplemento, não deixe de adicionar um prefixo ao nome do ícone com `ms-Icon--`.

    Para obter informações sobre tamanhos de fonte e cores disponíveis no Fabric Core, consulte [Typography](https://developer.microsoft.com/fluentui#/styles/web/typography) and the **Colors** table of contents at [Colors](https://developer.microsoft.com/fluentui#/styles/web/colors).

Exemplos são incluídos nos [Exemplos](#samples) posteriormente neste artigo.

## <a name="use-office-ui-fabric-js-components"></a>Usar Office UI Fabric JS

Os complementos com UIs não React também podem usar qualquer um dos muitos componentes do [Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js), incluindo botões, caixas de diálogo, seladores e muito mais. Consulte o readme do repo para obter instruções.

Exemplos são incluídos nos [Exemplos](#samples) posteriormente neste artigo.

## <a name="samples"></a>Exemplos

Os seguintes exemplos de complementos usam o Fabric Core e/ou Office UI Fabric componentes JS. Algumas dessas repos são arquivadas, o que significa que elas não estão mais sendo atualizadas com correções de bugs ou de segurança, mas você ainda pode usá-las para aprender a usar componentes do Fabric Core e da interface do usuário do Fabric.

- [Excel Add-in JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excel SalesLeads de complemento](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excel Tendências de despesas de woodgrove do add-in](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Office Exemplo de interface do usuário do Fabric do add-in](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlook Add-in GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPoint Add-in Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word Add-in Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
