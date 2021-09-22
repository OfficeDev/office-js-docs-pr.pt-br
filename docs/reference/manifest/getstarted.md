---
title: Elemento GetStarted no arquivo de manifesto
description: Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote.
ms.date: 09/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: c311e1bb5fbc2db265f430c8762216ad3a727107
ms.sourcegitcommit: a854a2fd2ad9f379a3ef712f307e0b1bb9b5b00d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2021
ms.locfileid: "59474340"
---
# <a name="getstarted-element"></a>Elemento GetStarted

Fornece informações usadas pelo texto explicante que aparece quando o complemento é instalado no Word, Excel, PowerPoint e OneNote. O elemento **GetStarted** é um elemento filho de [DesktopFormFactor](desktopformfactor.md). Se o **elemento GetStarted** for omitido, o explicativo usará os valores dos elementos [DisplayName](displayname.md) e [Description.](description.md)

## <a name="child-elements"></a>Elementos filho

| Elemento                       | Obrigatório | Descrição                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Sim      | Define onde um suplemento expõe a funcionalidade.     |
| [Descrição](#description)   | Sim      | Uma URL para um arquivo que contém funções JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Sim       | Uma URL para uma página que explica o suplemento em detalhes.   |

### <a name="title"></a>Título 

Obrigatório. O título usado para o início do texto explicativo. O **atributo resid** faz referência a uma ID válida no elemento **ShortStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

### <a name="description"></a>Descrição

Obrigatório. A descrição / conteúdo do corpo para o texto explicativo. O **atributo resid** faz referência a uma ID válida no elemento **LongStrings** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

### <a name="learnmoreurl"></a>LearnMoreUrl

Obrigatório. A URL para uma página onde o usuário pode saber mais sobre o suplemento. O **atributo resid** faz referência a uma ID válida no elemento **Urls** na seção [Recursos](resources.md) e não pode ter mais de 32 caracteres.

> [!NOTE]
> **LearnMoreUrl** atualmente não é processado em clientes do Word, Excel ou PowerPoint. Recomendamos que você adicione essa URL a todos os clientes para que a URL seja processada quando ficar disponível. 

## <a name="see-also"></a>Confira também

Os exemplos de código a seguir usam o **elemento GetStarted.**

* [Suplemento Web do Excel para manipular formatação de tabelas e gráficos](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit para um Suplemento do Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
