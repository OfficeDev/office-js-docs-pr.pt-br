---
title: Requisitos de suplementos do Outlook
description: Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 67aebd1fae19811797c07d33a5f6cac8907550f9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165852"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="85a1c-103">Requisitos de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="85a1c-103">Outlook add-in requirements</span></span>

<span data-ttu-id="85a1c-104">Há diversos requisitos para os servidores e clientes para que os Suplementos do Outlook possam carregar e funcionar de maneira apropriada.</span><span class="sxs-lookup"><span data-stu-id="85a1c-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="85a1c-105">Requisitos do cliente</span><span class="sxs-lookup"><span data-stu-id="85a1c-105">Client requirements</span></span>

- <span data-ttu-id="85a1c-106">O cliente deve ser um dos hosts suportados para suplementos do Outlook. Os clientes a seguir oferecem suporte para suplementos:</span><span class="sxs-lookup"><span data-stu-id="85a1c-106">The client must be one of the supported hosts for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="85a1c-107">Outlook 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="85a1c-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="85a1c-108">Outlook 2016 ou posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="85a1c-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="85a1c-109">Outlook no iOS</span><span class="sxs-lookup"><span data-stu-id="85a1c-109">Outlook on iOS</span></span>
   - <span data-ttu-id="85a1c-110">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="85a1c-110">Outlook on Android</span></span>
   - <span data-ttu-id="85a1c-111">Outlook na web para o Exchange 2016 ou posterior e Office 365</span><span class="sxs-lookup"><span data-stu-id="85a1c-111">Outlook on the web for Exchange 2016 or later and Office 365</span></span>
   - <span data-ttu-id="85a1c-112">Outlook na Web para Exchange 2013</span><span class="sxs-lookup"><span data-stu-id="85a1c-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="85a1c-113">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="85a1c-113">Outlook.com</span></span>

- <span data-ttu-id="85a1c-p101">O cliente deve ser conectado a um servidor do Exchange ou Office 365 usando uma conexão direta. Ao configurar o cliente, o usuário deve escolher um tipo de conta do **Exchange**, **Office 365** ou **Outlook.com**. Se o cliente estiver configurado para conectar-se com POP3 ou IMAP, os suplementos não serão carregados.</span><span class="sxs-lookup"><span data-stu-id="85a1c-p101">The client must be connected to an Exchange server or Office 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="85a1c-117">Requisitos de servidor de email</span><span class="sxs-lookup"><span data-stu-id="85a1c-117">Mail server requirements</span></span>

<span data-ttu-id="85a1c-p102">Se o usuário estiver conectado ao Office 365 ou ao Outlook.com, os requisitos do servidor de email já foram todos atendidos. No entanto, para os usuários conectados a instalações no local do Exchange Server, os seguintes requisitos se aplicam.</span><span class="sxs-lookup"><span data-stu-id="85a1c-p102">If the user is connected to Office 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="85a1c-120">O servidor deve ser o Exchange 2013 ou posterior.</span><span class="sxs-lookup"><span data-stu-id="85a1c-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="85a1c-121">Os Serviços Web do Exchange (EWS) devem estar habilitados e expostos à Internet.</span><span class="sxs-lookup"><span data-stu-id="85a1c-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="85a1c-122">Vários suplementos exigem o EWS para funcionar adequadamente.</span><span class="sxs-lookup"><span data-stu-id="85a1c-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="85a1c-123">O servidor deve ter um certificado de autenticação válido para que o servidor possa emitir tokens de identidade válidos.</span><span class="sxs-lookup"><span data-stu-id="85a1c-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="85a1c-124">Novas instalações do Servidor do Exchange incluem um certificado de autenticação padrão.</span><span class="sxs-lookup"><span data-stu-id="85a1c-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="85a1c-125">Para obter mais informações, confira [Certificados digitais e criptografia no Exchange 2016](/Exchange/architecture/client-access/certificates) e [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span><span class="sxs-lookup"><span data-stu-id="85a1c-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="85a1c-126">Para acessar os suplementos da [Appsource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), os servidores de acesso dos clientes devem conseguir se comunicar com a AppSource.</span><span class="sxs-lookup"><span data-stu-id="85a1c-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="85a1c-127">Requisitos de servidor de suplemento</span><span class="sxs-lookup"><span data-stu-id="85a1c-127">Add-in server requirements</span></span>

<span data-ttu-id="85a1c-p105">Os aquivos de suplemento (HTML, JavaScript, etc.) podem ser hospedados em qualquer plataforma de servidor Web desejada. O único requisito é que o servidor deve ser configurado para usar HTTPS e o cliente deve confiar no certificado SSL.</span><span class="sxs-lookup"><span data-stu-id="85a1c-p105">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="85a1c-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="85a1c-130">See also</span></span>

- [<span data-ttu-id="85a1c-131">Requisitos para a execução de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="85a1c-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="85a1c-132">Disponibilidade de host e plataforma para Suplementos do Office (seção do Outlook)</span><span class="sxs-lookup"><span data-stu-id="85a1c-132">Office Add-in host and platform availability (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="85a1c-133">Suporte ao conjunto de requisitos da API JavaScript do Outlook</span><span class="sxs-lookup"><span data-stu-id="85a1c-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)