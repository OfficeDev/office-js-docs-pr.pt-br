---
title: Elemento Hosts no arquivo de manifesto
description: Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718101"
---
# <a name="hosts-element"></a>Elemento Hosts

Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

Quando incluído no nó [VersionOverrides](versionoverrides.md), este elemento substitui o elemento **Hosts** na parte pai do manifesto. 

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
