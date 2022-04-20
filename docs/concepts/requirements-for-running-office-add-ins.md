---
title: Requisitos para a execução de Suplementos do Office
description: Saiba mais sobre os requisitos de cliente e servidor que um usuário final precisa para executar Office suplementos.
ms.date: 04/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9bc093b3e04dd1a67ba63bebbe2e44acf5137a07
ms.sourcegitcommit: 9795f671cacaa0a9b03431ecdfff996f690e30ed
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/20/2022
ms.locfileid: "64963488"
---
# <a name="requirements-for-running-office-add-ins"></a>Requisitos para a execução de Suplementos do Office

Este artigo descreve os requisitos de software e de dispositivo para execução de Suplementos do Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Para obter uma exibição de alto nível de onde Office suplementos têm suporte no momento, consulte Office disponibilidade de plataforma e aplicativo cliente para Office [suplementos](/javascript/api/requirement-sets).

## <a name="server-requirements"></a>Requisitos de servidor

Para poder instalar e executar qualquer Suplemento do Office, primeiro você precisa implantar os arquivos de manifesto e de página da Web para a interface de usuário e o código de seu suplemento para os locais de servidor apropriados.

Para todos os tipos de suplementos (suplementos de conteúdo, do Outlook e de painel de tarefas, além dos comandos de suplemento), você precisa implantar seus arquivos de página da Web do suplemento em um servidor Web ou em um serviço de hospedagem da Web, como o [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Ao desenvolver e depurar um suplemento no Visual Studio, este implanta e executa os arquivos de página da Web do suplemento localmente com o IIS Express, e não exige um servidor Web adicional.

Para suplementos de conteúdo e painel de tarefas, nos aplicativos cliente do Office com suporte – Excel, PowerPoint, Project ou Word – você também precisa de um catálogo de aplicativos no SharePoint para carregar o arquivo de manifesto [](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) XML do suplemento ou precisa implantar o suplemento usando Aplicativos Integrados[.](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)

Para testar e executar um suplemento do Outlook, a conta de email do Outlook do usuário deve residir no Exchange 2013 ou posterior, que está disponível por meio do Microsoft 365, Exchange Online ou por meio de uma instalação local. O usuário ou administrador instala os arquivos de manifesto para suplementos do Outlook nesse servidor.

> [!NOTE]
> Contas de email POP e IMAP no Outlook não são compatíveis com Suplementos do Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Requisitos de cliente: Windows para área de trabalho e tablet

O software a seguir é necessário para desenvolver um suplemento do Office para os clientes de área de trabalho do Office ou clientes Web com suporte que são executados em dispositivos desktop, laptop ou tablet baseados em Windows.

- Para computadores de mesa com Windows x86 e x64, e tablets como o Surface Pro:
  - A versão de 32 ou de 64 bits do Office 2013 ou uma versão posterior, em execução no Windows 7 ou em uma versão posterior.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013 ou uma versão posterior do cliente do Office, se você estiver testando ou executando um Suplemento do Office especificamente para um desses clientes de área de trabalho do Office. É possível instalar clientes de área de trabalho do Office localmente ou por meio do recurso Clique para Executar no computador cliente.

  Se você tiver uma assinatura Microsoft 365 válida e não tiver acesso ao cliente Office, poderá baixar e instalar a versão mais recente [do Office](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658).

- O Internet Explorer 11 ou o Microsoft Edge (dependendo das versões do Windows e do Office) devem estar instalados, mas não precisam ser o navegador padrão. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa componentes do navegador que fazem parte do Internet Explorer 11 ou do Microsoft Edge. Consulte [Navegadores usados pelos Suplementos do Office](browsers-used-by-office-web-add-ins.md) para obter mais detalhes.

  > [!NOTE]
  > A Configuração de Segurança Aprimorada da (ESC) do Internet Explorer deve ser desativada para os suplementos Web do Office funcionarem. Se estiver usando um computador Windows Server como cliente, ao desenvolver suplementos observe se a ESC está ativada por padrão no Windows Server.

- Um dos seguintes itens como o navegador padrão: Internet Explorer 11, Microsoft Edge em sua versão mais recente, Chrome, Firefox ou Safari (Mac OS).
- Um editor de HTML e JavaScript, como o Bloco de Notas, o [Visual Studio e Ferramentas de Desenvolvimento da Microsoft](https://www.visualstudio.com/features/office-tools-vs), ou uma ferramenta de desenvolvimento para Web de terceiros.

## <a name="client-requirements-os-x-desktop"></a>Requisitos de cliente: Computador com OS X

Outlook no Mac, que é distribuído como parte do Microsoft 365, dá suporte Outlook suplementos. Executar Outlook suplementos no Outlook no Mac tem os mesmos requisitos do Outlook no próprio Mac: o sistema operacional deve ser pelo menos o OS X v10.10 "Yosemite". Como o Outlook para Mac usa WebKit como um mecanismo de layout para processar as páginas do suplemento, não há qualquer dependência adicional de navegador.

Estas são as versões mínimas do cliente do Office para Mac que suporta suplementos do Office.

- Versão do Word 15.18 (160109)
- Versão do Excel 15.19 (160206)
- Versão do PowerPoint 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Requisitos de cliente: Suporte do navegador para clientes da Web do Office Online e SharePoint

Qualquer navegador, exceto o Internet Explorer, que dá suporte a ECMAScript 5.1, HTML5 e CSS3, como Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).

## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Requisitos do cliente: smartphone e tablet Windows não existentes

Especificamente para Outlook em execução em smartphones e dispositivos de tablet não Windows, o software a seguir é necessário para testar e executar Outlook suplementos.

| Aplicativo do Office | Dispositivo | Sistema operacional | Conta do Exchange | Navegador móvel |
|:-----|:-----|:-----|:-----|:-----|
|Outlook no Android|– Tablets Android<br>- Smartphones Android|– Android 4.4 KitKat ou posterior|Na atualização mais recente do Microsoft 365 Apps para Pequenos e Médios negócios ou Exchange Online|Navegador não aplicável. Use o aplicativo nativo para Android. <sup>1</sup>|
|Outlook no iOS|- iPad tablets<br>- iPhone smartphones|– iOS 11 ou posterior|Na atualização mais recente do Microsoft 365 Apps para Pequenos e Médios negócios ou Exchange Online|Navegador não aplicável. Use o aplicativo nativo para iOS. <sup>1</sup>|
|Outlook na Web (moderno)<sup>2</sup>|– iPad 2 ou posterior<br>– Tablets Android |– iOS 5 ou posterior<br>– Android 4.4 KitKat ou posterior|No Microsoft 365, Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>– Safari|
|Outlook na Web (clássico)|– iPhone 4 ou posterior<br>– iPad 2 ou posterior<br>- iPod Touch 4 ou posterior|– iOS 5 ou posterior|No local Exchange Server 2013 ou <sup>posterior3</sup>|– Safari|

> [!NOTE]
> <sup>1</sup> OWA para Android, OWA para iPad e OWA para iPhone aplicativos nativos foram [preteridos](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b).
>
> <sup>2</sup> Os Outlook na Web em iPhone e smartphones Android não são mais necessários ou estão disponíveis para teste Outlook suplementos.
>
> <sup>3</sup> Suplementos não têm suporte no Outlook android, no iOS e na Web móvel moderna com contas Exchange locais.

> [!TIP]
> É possível distinguir o Outlook clássico do moderno no navegador da Web, verificando sua barra de ferramentas da caixa de correio.
>
> **moderno**
>
> ![Captura de tela parcial da barra de ferramentas moderna do Outlook.](../images/outlook-on-the-web-new-toolbar.png)
>
> **clássico**
>
> ![Captura de tela parcial da barra de ferramentas clássica do Outlook.](../images/outlook-on-the-web-classic-toolbar.png)

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office](/javascript/api/requirement-sets)
- [Navegadores usados pelos Suplementos do Office](browsers-used-by-office-web-add-ins.md)
