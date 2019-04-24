---
title: Elemento GetStarted no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d9ebcba7881b388544eeb3e2c3028bff9bdcf9a6
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452078"
---
# <a name="getstarted-element"></a>Elemento GetStarted

Fornece informações usadas pelo balão que aparece quando o suplemento está instalado em hosts do Word, do Excel, do PowerPoint e do OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Elementos filho

| Elemento                       | Obrigatório | Descrição                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Sim      | Define onde um suplemento expõe a funcionalidade.     |
| [Descrição](#description)   | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Não       | Uma URL para uma página que explica o suplemento em detalhes.   |

### <a name="title"></a>Título 

Obrigatório. O título usado para o início do texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **ShortStrings** na seção [Recursos](resources.md).

### <a name="description"></a>Descrição

Obrigatório. A descrição / conteúdo do corpo para o texto explicativo. O atributo **resid** faz referência a uma identificação válida no elemento **LongStrings** na seção [Recursos](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O atributo **resid** faz referência a uma identificação válida no elemento **Urls** na seção [Recursos](resources.md).

> [!NOTE]
> **LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint. Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível. 

## <a name="see-also"></a>Confira também

Os exemplos de código a seguir utilizam o elemento **GetStarted**:

* [Suplemento Web do Excel para manipular formatação de tabelas e gráficos](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit para um Suplemento do Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
