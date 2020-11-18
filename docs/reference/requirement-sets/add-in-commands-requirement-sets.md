---
title: Conjuntos de requisitos dos comandos de suplemento
description: Visão geral dos conjuntos de requisitos de comandos de suplemento do Office.
ms.date: 11/01/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 08fcb5df0e614e9b9f3ec9479fc958cc79adf320
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087956"
---
# <a name="add-in-commands-requirement-sets"></a>Conjuntos de requisitos dos comandos de suplemento

Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Para saber mais, confira [Comandos de suplemento para Excel, Word e PowerPoint](../../design/add-in-commands.md) e [Comandos de suplemento para Outlook](../../outlook/add-in-commands-for-outlook.md).

A versão inicial dos comandos de suplemento não tem um conjunto de requisitos correspondente (ou seja, não há um conjunto de requisitos AddinCommands 1,0). A tabela a seguir lista os aplicativos cliente do Office que suportam a versão inicial do lançamento e as versões de compilação ou o número desses aplicativos.  

| Lançar   |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365)   |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Comandos de suplemento (versão inicial, nenhum conjunto de requisitos) | N/D | 16.0.4678.1000 *suportado somente no Outlook* | Versão 1809 (Build 10827.20150) ou posterior |Versão 1603 (Build 6769.0000) ou posterior | N/D | 15.33 ou posterior| Janeiro de 2016 |

O conjunto de requisitos **1,1** de comandos de suplemento introduz a capacidade de [abrir o painel de tarefas com documentos](../../develop/automatically-open-a-task-pane-with-a-document.md).

Os comandos de suplemento **1,3** conjunto de requisitos introduz a marcação de manifesto que permite que um suplemento Personalize o posicionamento de uma guia personalizada na faixa de opções do Office e insira controles de faixa de opções do Office internos em grupos de controle personalizados.

A tabela a seguir lista os conjuntos de requisitos de comandos de suplemento, os aplicativos cliente do Office que dão suporte a esse conjunto de requisitos e os números de versão ou de compilação do aplicativo do Office.

|  Conjunto de requisitos  |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office no Windows<br>(conectado a uma assinatura do Microsoft 365)   |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac<br>(conectado a uma assinatura do Microsoft 365)  | Office na Web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1,3  | N/D | N/D  | N/D | em breve | N/D | em breve | Novembro de 2020 |
| AddInCommands 1.1  | N/D | 16.0.4678.1000 *suportado somente no Outlook*  | Versão 1809 (Build 10827.20150) ou posterior | Versão 1705 (Build 8121.1000) ou posterior | N/D | 15.34 ou posterior\*| Maio de 2017 |

>\*O método [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) retornará `false` erroneamente para as versões 16.9 &ndash; 16.14 (incluindo), mas o conjunto de requisitos *é* suportado nessas versões.

> [!IMPORTANT]
> AddinCommands 1,3 está em versão prévia e *só está disponível no PowerPoint na Web*. Recomendamos que você experimente a marcação apenas em ambientes de teste e desenvolvimento. Não use a marcação de visualização em um ambiente de produção ou em documentos de negócios críticos.

## <a name="office-versions-and-build-numbers"></a>Versões do Office e números de build

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
