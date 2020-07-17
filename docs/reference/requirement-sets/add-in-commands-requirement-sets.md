---
title: Conjuntos de requisitos dos comandos de suplemento
description: Visão geral dos conjuntos de requisitos de comandos de suplemento do Office.
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 14952f97e3f03b83c15d0a4d51183fac3ec80af3
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159084"
---
# <a name="add-in-commands-requirement-sets"></a>Conjuntos de requisitos dos comandos de suplemento

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Para saber mais, confira [Comandos de suplemento para Excel, Word e PowerPoint](../../design/add-in-commands.md) e [Comandos de suplemento para Outlook](../../outlook/add-in-commands-for-outlook.md).

A versão inicial dos comandos do suplemento não tem um conjunto de requisitos correspondente (ou seja, não há um conjunto de requisitos 1.0 de AddInCommands). A tabela a seguir lista os aplicativos de host do Office que oferecem suporte à versão de lançamento inicial e os números de versão ou de build dos aplicativos.  

| Lançar   |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365)   |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Comandos de suplemento (versão inicial, nenhum conjunto de requisitos) | N/D | 16.0.4678.1000 *suportado somente no Outlook* | Versão 1809 (Build 10827.20150) ou posterior |Versão 1603 (Build 6769.0000) ou posterior | N/D | 15.33 ou posterior| Janeiro de 2016 |

O conjunto de requisitos 1.1 dos comandos do suplemento introduz a capacidade de [abrir automaticamente um painel de tarefas com documentos](../../develop/automatically-open-a-task-pane-with-a-document.md).

A tabela a seguir lista o conjunto de requisitos 1.1 dos comandos do suplemento, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos e os números de versão ou de build dos aplicativos do Office.

|  Conjunto de requisitos  |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365)   |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |  
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
