---
title: Conjuntos de requisitos dos comandos de suplemento
description: Visão geral dos conjuntos de requisitos de comandos de suplemento do Office
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d1904d092988a445be3e481123eecbad39097764
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094411"
---
# <a name="add-in-commands-requirement-sets"></a>Conjuntos de requisitos dos comandos de suplemento

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| Lançar   |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado à assinatura do Microsoft 365)   |  Office no iPad<br>(conectado à assinatura do Microsoft 365)  |  Office no Mac<br>(conectado à assinatura do Microsoft 365)  | Office na Web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Comandos de suplemento (versão inicial, nenhum conjunto de requisitos) | N/D | 16.0.4678.1000 *suportado somente no Outlook* | Versão 1809 (Build 10827.20150) ou posterior |Versão 1603 (Build 6769.0000) ou posterior | N/D | 15.33 ou posterior| Janeiro de 2016 |

O conjunto de requisitos 1.1 dos comandos do suplemento introduz a capacidade de [abrir automaticamente um painel de tarefas com documentos](../../develop/automatically-open-a-task-pane-with-a-document.md).

A tabela a seguir lista o conjunto de requisitos 1.1 dos comandos do suplemento, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado à assinatura do Microsoft 365)   |  Office no iPad<br>(conectado à assinatura do Microsoft 365)  |  Office no Mac<br>(conectado à assinatura do Microsoft 365)  | Office na Web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | N/D | 16.0.4678.1000 *suportado somente no Outlook*  | Versão 1809 (Build 10827.20150) ou posterior | Versão 1705 (Build 8121.1000) ou posterior | N/D | 15.34 ou posterior\*| Maio de 2017 |

>\*O método [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) retornará `false` erroneamente para as versões 16.9 &ndash; 16.14 (incluindo), mas o conjunto de requisitos * é *suportado nessas versões.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de hosts do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
