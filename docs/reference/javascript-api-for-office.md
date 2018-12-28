---
title: API JavaScript para Office
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1f57ec9e4420a17ef0997d8d293c484887d5d79
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432771"
---
# <a name="javascript-api-for-office"></a>API JavaScript para Office

A API JavaScript para Office permite que você crie aplicativos Web que interajam com os modelos de objeto em aplicativos host do Office. Seu aplicativo fará referência à biblioteca office.js, que é um carregador de script. A biblioteca office.js carrega os modelos de objeto que são aplicáveis ao aplicativo do Office em execução no suplemento. Você pode usar os seguintes modelos de objeto JavaScript:

- **APIs comuns**: APIs introduzidas com o **Office 2013**. Elas são carregadas em **todos os aplicativos host do Office** e conectam seu aplicativo de suplemento com o aplicativo cliente do Office. O modelo de objeto contém APIs específicas dos clientes do Office e APIs aplicáveis a vários aplicativos host de cliente do Office. Todo esse conteúdo está em **API compartilhada**. 

  O **Outlook** também usa a sintaxe comum de API. Todo o conteúdo sob o alias Office contém objetos que você pode usar para gravar scripts que interagem com o conteúdo em documentos, planilhas, apresentações, itens de email e projetos do Office a partir de seus suplementos do Office. Você deve usar essas APIs comuns se o seu suplemento tiver como meta o Office 2013 e posterior. Este modelo de objeto usa retornos de chamada.

- **APIs específicas de host**: APIs introduzidas com o **Office 2016**. Este modelo de objeto fornece objetos fortemente tipados e específicos do host que correspondem aos objetos familiares exibidos quando você usa os clientes do Office, e representa o futuro das APIs JavaScript para Office. No momento, as APIs específicas do host incluem a API JavaScript do Word e a API JavaScript do Excel.

## <a name="supported-host-applications"></a>Aplicativos host compatíveis

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API compartilhada](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> O [PowerPoint e o Project](requirement-sets/powerpoint-and-project-note.md) são compatíveis com suplementos feitos com a API JavaScript. No entanto, eles atualmente não possuem APIs específicas do host. Você interage com esses hosts por meio da API compartilhada.

Saiba mais sobre [hosts compatíveis e outros requisitos](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

## <a name="open-api-specifications"></a>Especificações abertas da API

À medida que criamos e desenvolvemos novas APIs para suplementos do Office, nós as disponibilizamos em nossa página [Especificações abertas da API](openspec.md) a fim de obter os seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.

## <a name="see-also"></a>Confira também

- [Referência da API JavaScript do Office](https://docs.microsoft.com/javascript/api/overview/office?view=office-js)