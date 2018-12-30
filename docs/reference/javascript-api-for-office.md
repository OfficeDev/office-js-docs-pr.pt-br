---
title: API JavaScript para Office
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 87ad98f8233e4ff6fb2fe15d09daff6b7b422b08
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457709"
---
# <a name="javascript-api-for-office"></a>API JavaScript para Office

A API JavaScript para Office permite que você crie aplicativos Web que interajam com os modelos de objeto em aplicativos host do Office. Seu aplicativo fará referência à biblioteca office.js, que é um carregador de script. A biblioteca office.js carrega os modelos de objeto que são aplicáveis ao aplicativo do Office em execução no suplemento. Você pode usar os seguintes modelos de objeto JavaScript:

- **APIs comuns**: APIs introduzidas com o **Office 2013**. Elas são carregadas em **todos os aplicativos host do Office** e conectam seu aplicativo de suplemento com o aplicativo cliente do Office. O modelo de objeto contém APIs específicas aos clientes do Office e APIs aplicáveis a vários aplicativos host de clientes do Office. Todo esse conteúdo está na **API Comum**. Este modelo de objeto usa retornos de chamada. 

  O **Outlook** também usa a sintaxe da API Comum. Todo o conteúdo sob o alias Office contém objetos que você pode usar para gravar scripts que interagem com o conteúdo em documentos, planilhas, apresentações, itens de email e projetos do Office a partir de seus suplementos do Office. Você deve usar essas APIs Comuns se o seu suplemento servir para o Office 2013 e versões posteriores. Este modelo de objeto usa retornos de chamada.

- **APIs específicas de host**: APIs introduzidas com o **Office 2016**. Este modelo de objeto fornece objetos fortemente tipados e específicos do host que correspondem aos objetos familiares exibidos quando você usa os clientes do Office, e representa o futuro das APIs JavaScript para Office. No momento, as APIs específicas do host incluem a API JavaScript do Word e a API JavaScript do Excel.

## <a name="supported-host-applications"></a>Aplicativos host compatíveis

- [Excel](overview/excel-add-ins-reference-overview.md)
- [OneNote](overview/onenote-add-ins-javascript-reference.md)
- [Outlook](requirement-sets/outlook-api-requirement-sets.md)
- [Visio](overview/visio-javascript-reference-overview.md)
- [Word](overview/word-add-ins-reference-overview.md)
- [API Comum](requirement-sets/office-add-in-requirement-sets.md)

> [!NOTE] 
> O [PowerPoint e o Project](requirement-sets/powerpoint-and-project-note.md) são compatíveis com suplementos feitos com a API JavaScript. No entanto, eles atualmente não possuem APIs específicas do host. Você pode interagir com esses hosts por meio da API Comum.

Saiba mais sobre [hosts suportados e outros requisitos](../concepts/requirements-for-running-office-add-ins.md).

## <a name="open-api-specifications"></a>Especificações abertas da API

À medida que criamos e desenvolvemos novas APIs para suplementos do Office, nós as disponibilizamos em nossa página [Especificações abertas da API](openspec.md) a fim de obter os seus comentários. Descubra quais novos recursos estão no pipeline e forneça comentários sobre nossas especificações de design.

## <a name="see-also"></a>Confira também

- [Referência da API JavaScript do Office](https://docs.microsoft.com/javascript/api/overview/office)