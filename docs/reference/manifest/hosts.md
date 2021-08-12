---
title: Elemento Hosts no arquivo de manifesto
description: Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: c89a0154b2dbbc9b07a10493401ff761d48b955d7538eb14a825591d2b12607d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57083799"
---
# <a name="hosts-element"></a>Elemento Hosts

Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

Quando incluído no nó [VersionOverrides](versionoverrides.md), este elemento substitui o elemento **Hosts** na parte pai do manifesto. 

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
