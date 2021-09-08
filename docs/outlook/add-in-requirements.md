---
title: Requisitos de suplementos do Outlook
description: Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 6062073d44a412d67961f806677cd60701bbdb9b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936255"
---
# <a name="outlook-add-in-requirements"></a>Requisitos de suplementos do Outlook

Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.

## <a name="client-requirements"></a>Requisitos do cliente

- O cliente deve ser um dos aplicativos com suporte para suplementos do Outlook. Os clientes a seguir dão suporte a suplementos.

  - Outlook 2013 ou posterior no Windows
  - Outlook 2016 ou posterior no Mac
  - Outlook no iOS
  - Outlook no Android
  - Outlook na Web para o Exchange 2016 ou posterior
  - Outlook na Web para Exchange 2013
  - Outlook.com

- O cliente deve estar conectado a um servidor Exchange ou Microsoft 365 usando uma conexão direta. Ao configurar o cliente, o usuário deve escolher um tipo de conta do **Exchange**, **Office** ou **Outlook.com**. Se o cliente estiver configurado para se conectar com POP3 ou IMAP, os suplementos não serão carregados.

## <a name="mail-server-requirements"></a>Requisitos de servidor de email

Se o usuário estiver conectado ao Microsoft 365 ou ao Outlook.com, os requisitos do servidor de email já foram todos atendidos. No entanto, para os usuários conectados a instalações locais do Exchange Server, os seguintes requisitos se aplicam.

- O servidor deve ser o Exchange 2013 ou posterior.
- Os Serviços Web do Exchange (EWS) devem estar habilitados e expostos à Internet. Vários suplementos exigem o EWS para funcionar adequadamente.
- O servidor deve ter um certificado de autenticação válido para que o servidor possa emitir tokens de identidade válidos. Novas instalações do Servidor do Exchange incluem um certificado de autenticação padrão. Para obter mais informações, confira [Certificados digitais e criptografia no Exchange 2016](/Exchange/architecture/client-access/certificates) e [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- Para acessar os suplementos da [Appsource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), os servidores de acesso dos clientes devem conseguir se comunicar com a AppSource.

## <a name="add-in-server-requirements"></a>Requisitos de servidor de suplemento

Os aquivos de suplemento (HTML, JavaScript, etc.) podem ser hospedados em qualquer plataforma de servidor Web desejada. O único requisito é que o servidor deve ser configurado para usar HTTPS e o cliente deve confiar no certificado SSL.

## <a name="see-also"></a>Confira também

- [Requisitos para a execução de suplementos do Office](../concepts/requirements-for-running-office-add-ins.md)
- [Disponibilidade de aplicativos e plataformas de cliente do Office para Suplementos do Office (seção do Outlook)](../overview/office-add-in-availability.md#outlook)
- [Suporte ao conjunto de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
