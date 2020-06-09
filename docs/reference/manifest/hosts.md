---
title: Elemento Hosts no arquivo de manifesto
description: Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611803"
---
# <a name="hosts-element"></a>Elemento Hosts

Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

Quando incluído no nó [VersionOverrides](versionoverrides.md), este elemento substitui o elemento **Hosts** na parte pai do manifesto. 

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
