---
title: Requisitos de suplementos do Outlook
description: Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 700e0efd2ab2655de61d37d42038fa2c15a99cb4
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093991"
---
# <a name="outlook-add-in-requirements"></a>Requisitos de suplementos do Outlook

Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.

## <a name="client-requirements"></a>Requisitos do cliente

- O cliente deve ser um dos hosts suportados para suplementos do Outlook. Os clientes a seguir oferecem suporte para suplementos:

   - Outlook 2013 ou posterior no Windows
   - Outlook 2016 ou posterior no Mac
   - Outlook no iOS
   - Outlook no Android
   - Outlook na web para o Exchange 2016 ou posterior e Office 365
   - Outlook na Web para Exchange 2013
   - Outlook.com

- The client must be connected to an Exchange server or Microsoft 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.

## <a name="mail-server-requirements"></a>Requisitos de servidor de email

If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.

- O servidor deve ser o Exchange 2013 ou posterior.
- Os Serviços Web do Exchange (EWS) devem estar habilitados e expostos à Internet. Vários suplementos exigem o EWS para funcionar adequadamente.
- O servidor deve ter um certificado de autenticação válido para que o servidor possa emitir tokens de identidade válidos. Novas instalações do Servidor do Exchange incluem um certificado de autenticação padrão. Para obter mais informações, confira [Certificados digitais e criptografia no Exchange 2016](/Exchange/architecture/client-access/certificates) e [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).
- Para acessar os suplementos da [Appsource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), os servidores de acesso dos clientes devem conseguir se comunicar com a AppSource.

## <a name="add-in-server-requirements"></a>Requisitos de servidor de suplemento

Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.

## <a name="see-also"></a>Confira também

- [Requisitos para a execução de suplementos do Office](../concepts/requirements-for-running-office-add-ins.md)
- [Disponibilidade de host e plataforma para Suplementos do Office (seção do Outlook)](../overview/office-add-in-availability.md#outlook)
- [Suporte ao conjunto de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
