---
title: Conjuntos de requisitos dos comandos de suplemento
description: Visão geral Office conjuntos de requisitos de comandos de complemento.
ms.date: 02/15/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: e35a36c10f9ca275d5dd969a3592df42a5e1000a
ms.sourcegitcommit: 789545a81bd61ec2e7adef2bc24c06b5be113b00
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/18/2022
ms.locfileid: "62892542"
---
# <a name="add-in-commands-requirement-sets"></a>Conjuntos de requisitos dos comandos de suplemento

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

Comandos de suplemento são elementos de interface do usuário que estendem a interface do usuário do Office e iniciam ações no suplemento. Você pode usar comandos de suplemento para adicionar um botão à faixa de opções ou um item a um menu de contexto. Para saber mais, confira [Comandos de suplemento para Excel, Word e PowerPoint](../../design/add-in-commands.md) e [Comandos de suplemento para Outlook](../../outlook/add-in-commands-for-outlook.md).

> [!NOTE]
> Outlook os complementos suportam comandos de add-in, mas as APIs e os elementos de manifesto que habilitam comandos de complemento no Outlook estão no conjunto de requisitos de Caixa [de Correio 1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md). Os conjuntos de requisitos AddinCommands não são aplicáveis a Outlook.

A versão inicial dos comandos de complemento não tem um conjunto de requisitos correspondente (ou seja, não há um conjunto de requisitos AddinCommands 1.0). A tabela a seguir lista os aplicativos Office cliente que suportam a versão de versão inicial e as versões de com build ou número desses aplicativos.  

| Lançar   |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) | Office 2021 no Windows<br>(compra avulsa) | Office no Windows<br>(assinatura)   |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Comandos de suplemento (versão inicial, nenhum conjunto de requisitos) | N/D | N/D | Versão 1809 (Build 10827.20150) ou posterior| 16.0.14326.20454 ou posterior |Versão 1603 (Build 6769.0000) ou posterior | N/D | 15.33 ou posterior| Janeiro de 2016 |

O conjunto de requisitos **de comandos de complemento 1.1** introduz a capacidade de [abrir automaticamente um painel de tarefas com documentos](../../develop/automatically-open-a-task-pane-with-a-document.md).

O conjunto de requisitos de comandos do add-in **1.3** introduz a marcação de manifesto que permite que um complemento personalize o posicionamento de uma guia personalizada na faixa de opções do Office e insira controles de faixa de opções Office internos em grupos de controle personalizados.

A tabela a seguir lista os conjuntos de requisitos de comandos de Office, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  |  Office 2013 no Windows<br>(compra avulsa) | Office 2016 no Windows<br>(compra avulsa) | Office 2019 no Windows<br>(compra avulsa) |  Office 2021 no Windows<br>(compra avulsa) | Office no Windows<br>(assinatura)   |  Office no iPad<br>(assinatura)  |  Office no Mac<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | N/D | N/D | N/D | N/D | Sem suporte | N/D | Sem suporte | Novembro de 2020 |
| AddInCommands 1.1  | N/D | N/D  | Versão 1809 (Build 10827.20150) ou posterior&dagger; | 16.0.14326.20454 ou posterior&dagger; | Versão 1705 (Build 8121.1000) ou posterior&dagger; | N/D | 15.34 ou posterior&dagger;\*| Maio de 2017 |

\*O método [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) retornará `false` erroneamente para as versões 16.9 &ndash; 16.14 (incluindo), mas o conjunto de requisitos *é* suportado nessas versões.

&dagger;OneNote é suportado somente em Office na Web.

> [!IMPORTANT]
> AddinCommands 1.3 está na visualização e *está disponível apenas no PowerPoint na Web*. Recomendamos que você experimente apenas a marcação em ambientes de teste e desenvolvimento. Não use a marcação de visualização em um ambiente de produção ou em documentos críticos para os negócios.

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
