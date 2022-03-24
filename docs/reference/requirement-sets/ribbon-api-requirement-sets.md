---
title: Conjuntos de requisitos comuns da API
description: Especifica quais plataformas Office e builds suportam as APIs dinâmicas da faixa de opções.
ms.date: 03/12/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: e399fe14626da2abd688b0e486454908ce1e9f91
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746689"
---
# <a name="ribbon-api-requirement-sets"></a>Conjuntos de requisitos comuns da API

Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).

O conjunto de APIs da Faixa de Opções dá suporte ao controle programático de quando os comandos de complemento personalizados (ou seja, botões de faixa de opções personalizados e itens de menu) são habilitados e desabilitados e quando as guias contextuais aparecem na faixa de opções.

> [!NOTE]
> Os conjuntos de requisitos RibbonApi são suportados somente em complementos do painel de tarefas.

Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da API da Faixa de Opções, os aplicativos cliente Office que suportam esse conjunto de requisitos e os números de com build ou versão do aplicativo Office.

|  Conjunto de requisitos  | Office 2021 ou posterior no Windows\*<br>(compra avulsa) | Office no Windows\*<br>(conectado a uma assinatura do Microsoft 365) |  Office no iPad<br>(conectado a uma assinatura do Microsoft 365)  |  Office no Mac\*<br>(ambas as assinaturas<br> e compra única Office no Mac 2019 e posterior)   | Office na Web\*  |  Servidor do Office Online  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.2  | Build 16.0.14326.20454 ou posterior | 2102 (build 13801.20294) | N/D | 16.53.806.0 | Maio de 2021 | N/D|
| RibbonApi 1.1  | Build 16.0.14326.20454 ou posterior | Consulte suporte<br>seção abaixo | N/D | 16.38 | Novembro de 2020 | N/D|

> **&#42;** A API da Faixa de Opções é suportada somente Excel.

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>Suporte para a versão 1.1 no Office no Windows (assinatura)

A versão 1.1 do conjunto de requisitos RibbonApi é suportada no Canal do Consumidor versão 2006 (build 13001.20498 ou superior). Para Office no Windows o recurso também é suportado nas builds do Canal Semi-Annual e do Canal mensal Enterprise disponíveis em 14 de julho de 2020 ou posterior. As builds mínimas com suporte para cada canal são as seguinte:  

|Canal | Versão | Build|
|:-----|:-----|:-----|
|Canal Atual | 2006 ou superior | 20266.20266 ou superior|
|Canal Empresarial Mensal | 2005 ou superior | 12827.20538 ou superior|
|Canal Empresarial Mensal | 2004 | 12730.20602 ou superior|
|Canal Empresarial Semestral | 2002 ou superior | 12527.20880 ou superior|

## <a name="more-information"></a>Mais informações

Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:

- [Números de versão e com build de versões de canal de atualização para Microsoft 365 clientes](/officeupdates/update-history-microsoft365-apps-by-date)
- [Qual versão do Office estou usando?](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Onde você pode encontrar a versão e o número de com build de um aplicativo Microsoft 365 cliente](/officeupdates/update-history-microsoft365-apps-by-date)
- [Visão geral sobre o Servidor do Office Online](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Conjuntos de requisitos da API comum do Office

Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API da Faixa de Opções 1.1

A API da Faixa de Opções 1.1 é a primeira versão da API. Para obter detalhes sobre a API, consulte o [tópico Office.ribbon](/javascript/api/office/office.ribbon) reference.

## <a name="ribbon-api-12"></a>API da Faixa de Opções 1.2

A API da Faixa de Opções 1.2 adiciona suporte a guias contextuais. Para obter mais informações, confira [Criar guias contextuais personalizadas em Suplementos do Office](../../design/contextual-tabs.md).

> [!NOTE]
> O **conjunto de requisitos RibbonApi 1.2** ainda não tem suporte no manifesto, portanto, você não deve especificá-lo na seção do `<Requirements>` manifesto.

## <a name="see-also"></a>Confira também

- [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md)
- [Especificar requisitos da API e de aplicativos do Office](../../develop/specify-office-hosts-and-api-requirements.md)
- [Manifesto XML dos Suplementos do Office](../../develop/add-in-manifests.md)
