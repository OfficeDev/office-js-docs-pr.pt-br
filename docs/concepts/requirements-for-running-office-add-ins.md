---
title: Requisitos para a execu??o de Suplementos do Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: a4859af73d8e9cf041990a3533894b24f1cbde6f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="requirements-for-running-office-add-ins"></a>Requisitos para a execu??o de Suplementos do Office

Este artigo descreve os requisitos de software e de dispositivo para execu??o de Suplementos do Office.

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

Confira uma vis?o avan?ada da compatibilidade atual dos suplementos do Office no momento na p?gina [Disponibilidade de hosts e plataformas de suplementos do Office](../overview/office-add-in-availability.md).

## <a name="server-requirements"></a>Requisitos de servidor

Para poder instalar e executar qualquer Suplemento do Office, primeiro voc? precisa implantar os arquivos de manifesto e de p?gina da Web para a interface de usu?rio e o c?digo de seu suplemento para os locais de servidor apropriados.

Para todos os tipos de suplementos (suplementos de conte?do, do Outlook e de painel de tarefas, al?m dos comandos de suplemento), voc? precisa implantar seus arquivos de p?gina da Web do suplemento em um servidor Web ou em um servi?o de hospedagem da Web, como o [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Ao desenvolver e depurar um suplemento no Visual Studio, este implanta e executa os arquivos de p?gina da Web do suplemento localmente com o IIS Express, e n?o exige um servidor Web adicional. 

Para suplementos de conte?do e de painel de tarefas, nos aplicativos host do Office compat?veis (aplicativos Web do Access, Word, Excel, PowerPoint ou Project) voc? tamb?m precisa de um [cat?logo de suplementos](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) no SharePoint para carregar os arquivo de manifesto XML do suplemento.

Para testar e executar um suplemento do Outlook, a conta de email do Outlook do usu?rio deve residir no Exchange 2013 ou posterior, que est? dispon?vel pelo Office 365, Exchange Online ou por meio de uma instala??o local. O usu?rio ou administrador instala os arquivos de manifesto para suplementos do Outlook nesse servidor.

> [!NOTE]
> Contas de email POP e IMAP no Outlook n?o s?o compat?veis com Suplementos do Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Requisitos de cliente: Windows para ?rea de trabalho e tablet

O software a seguir ? necess?rio para o desenvolvimento de um Suplemento do Office para os clientes de ?rea de trabalho ou da Web do Office compat?veis que s?o executados em dispositivos de ?rea de trabalho, laptop ou tablet baseados em Windows:


- Para computadores de mesa com Windows x86 e x64, e tablets como o Surface Pro:
    - A vers?o de 32 ou de 64 bits do Office 2013 ou uma vers?o posterior, em execu??o no Windows 7 ou em uma vers?o posterior.
    - Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013 ou uma vers?o posterior do cliente do Office, se voc? estiver testando ou executando um Suplemento do Office especificamente para um desses clientes de ?rea de trabalho do Office. ? poss?vel instalar clientes de ?rea de trabalho do Office localmente ou por meio do recurso Clique para Executar no computador cliente.
    
  Se voc? tem uma assinatura v?lida do Office 365 e n?o tem acesso ao Office 2013, voc? pode baix?-lo por meio de um dos links CDN:       
    - [Office 2013 para Empresas (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
    - [Office 2013 para Casa (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- Internet Explorer 11 ou posterior, que deve estar instalado, mas n?o precisa ser o navegador padr?o. Para oferecer suporte aos Suplementos do Office, o cliente do Office que atua como host usa os componentes do navegador que fazem parte do Internet Explorer 11 ou posterior.
- Um dos navegadores seguintes como o padr?o: Internet Explorer 11 ou posterior, ou a vers?o mais recente do Microsoft Edge, Chrome, Firefox ou Safari (Mac OS).
- Um editor de HTML e JavaScript, como o Bloco de Notas, o [Visual Studio e Ferramentas de Desenvolvimento da Microsoft](https://www.visualstudio.com/features/office-tools-vs), ou uma ferramenta de desenvolvimento para Web de terceiros.

## <a name="client-requirements-os-x-desktop"></a>Requisitos de cliente: Computador com OS X

Outlook para Mac, que ? distribu?do como parte do Office 365 e oferece suporte a suplementos do Outlook. A execu??o de suplementos do Outlook no Outlook para Mac tem os mesmos requisitos que o pr?prio Outlook para Mac: o sistema operacional deve ser pelo menos o OS X v10.10 "Yosemite". Como o Outlook para Mac usa WebKit como um mecanismo de layout para processar as p?ginas do suplemento, n?o h? qualquer depend?ncia adicional de navegador.

Estas s?o as vers?es m?nimas do cliente do Office para Mac que oferecem suporte a suplementos do Office:

- Word para Mac vers?o 15.18 (160109) 
- Excel para Mac vers?o 15.19 (160206) 
- PowerPoint para Mac vers?o 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>Requisitos de cliente: Suporte do navegador para clientes da Web do Office Online e SharePoint

Qualquer navegador compat?vel com ECMAScript 5.1, HTML5 e CSS3, como o Internet Explorer 11 ou posterior, ou a vers?o mais recente do Microsoft Edge, do Chrome, do Firefox ou do Safari (Mac OS).


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Requisitos do cliente: smartphones e tablets sem Windows

Especificamente para o OWA para Dispositivos e o Outlook Web App em execu??o em um navegador em smartphones e tablets sem Windows, os softwares a seguir s?o necess?rios para testar e executar suplementos do Outlook.


| Aplicativo host | Dispositivo | Sistema operacional | Conta do Exchange | Navegador m?vel |
|:-----|:-----|:-----|:-----|:-----|
|OWA para Android|Smartphones Android. Tecnicamente, os dispositivos considerados "pequenos" ou "normais" pelo [SO Android](https://developer.android.com/guide/practices/screens_support.html).|Android 4.4 KitKat ou posterior|Atualiza??o mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para Android, navegador n?o aplic?vel|
|OWA para iPad|iPad 2 ou posterior|iOS 6 ou posterior|Atualiza??o mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para iOS, navegador n?o aplic?vel|
|OWA para iPhone|iPhone 4S ou posterior|iOS 6 ou posterior|Atualiza??o mais recente do Office 365 para empresas ou do Exchange Online|Suplemento nativo para iOS, navegador n?o aplic?vel|
|Outlook Web App|iPhone 4 ou posterior, iPad 2 ou posterior, iPod Touch 4 ou posterior|iOS 5 ou posterior|Office 365, Exchange Online ou Exchange Server 2013 local ou posteriores|Safari|


## <a name="see-also"></a>Veja tamb?m

- [Vis?o geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Disponibilidade de host e plataforma para suplementos do Office](../overview/office-add-in-availability.md)
