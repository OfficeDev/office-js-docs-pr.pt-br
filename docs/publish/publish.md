---
title: Implantar e publicar seu suplemento do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: d8264667306dcdac2e9d5e5d6e6607a2a2100546
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="deploy-and-publish-your-office-add-in"></a>Implantar e publicar seu suplemento do Office

Voc? pode usar um dos v?rios m?todos para implantar o suplemento do Office para teste ou distribui??o aos usu?rios.

|**M?todo**|**Uso...**|
|:---------|:------------|
|[Sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Usado como parte do processo de desenvolvimento para testar o suplemento em execu??o no Windows, Office Online, iPad ou Mac.|
|[Implanta??o Centralizada](centralized-deployment.md)|Em uma implanta??o h?brida ou de nuvem para distribuir seu suplemento aos usu?rios na sua organiza??o usando o centro de administra??o do Office 365.|
|[Cat?logo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Usado para distribuir o suplemento aos usu?rios da organiza??o em um ambiente local.|
|[AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store)|Usado para distribuir o suplemento publicamente aos usu?rios.|
|[Servidor Exchange](#outlook-add-in-deployment)|Usado para distribuir suplementos do Outlook aos usu?rios em um ambiente local ou online.|
|[Compartilhamento de rede](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|No computador do Windows em uma rede na qual voc? deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que voc? deseja usar como seu cat?logo de pasta compartilhada.|

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experi?ncia do Office depois de cri?-lo, verifique se voc? est? em conformidade com as [Pol?ticas de valida??o do AppSource](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). Por exemplo, para passar na valida??o, seu suplemento deve funcionar em todas as plataformas com suporte aos m?todos que voc? definir (para mais informa??es, confira a [se??o 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [P?gina de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).

## <a name="deployment-options-by-office-host"></a>Op??es de implanta??o pelo host do Office

As op??es de implanta??o dispon?veis dependem do host do Office que voc? pretende usar e do tipo de suplemento que voc? pretende criar.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Op??es de implanta??o de suplementos para Word, Excel e PowerPoint

| Ponto de extens?o | Sideloading | Centro de administra??o do Office 365 |AppSource| Cat?logo do SharePoint\*  |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| Conte?do         | X           | X                       | X          | X                    |
| Painel de tarefas       | X           | X                       | X          | X                    |
| Comando           | X           | X                       | X          |                      |

* Os cat?logos do SharePoint n?o s?o compat?veis com o Office 2016 para Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Op??es de implanta??o para suplementos do Outlook

| Ponto de extens?o | Sideloading | Servidor Exchange | AppSource |
|:----------------|:-----------:|:---------------:|:------------:|
| Aplicativo de email        | X           | X               | X            |
| Comando         | X           | X               | X            |

## <a name="deployment-methods"></a>M?todos de implanta??o

As se??es a seguir fornecem informa??es adicionais sobre os m?todos de implanta??o mais comumente usados para distribuir suplementos do Office para usu?rios da organiza??o.

Saiba mais sobre como os usu?rios finais podem adquirir, inserir e executar suplementos em [Come?ar a usar seu Suplemento do Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>Implanta??o Centralizada por meio do centro de administra??o do Office 365 

No Centro de administra??o do Office 365, ? mais f?cil para o administrador implantar Suplementos do Office para usu?rios e grupos dentro da organiza??o. Os suplementos implantados por meio do Centro de administra??o ficam dispon?veis imediatamente para os usu?rios nos aplicativos do Office, sem a necessidade de configura??o do cliente. ? poss?vel usar a Implanta??o Centralizada para implantar suplementos internos, al?m de suplementos fornecidos por ISVs.

Confira mais informa??es em [Publicar Suplementos do Office usando a Implanta??o Centralizada por meio do Centro de Administra??o do Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>Implanta??o de cat?logo do SharePoint

O cat?logo de suplementos do SharePoint ? uma cole??o de sites especial que voc? pode criar para hospedar suplementos dos aplicativos Word, Excel e PowerPoint. Como os cat?logos do SharePoint n?o oferecem suporte para os novos recursos de suplemento implementados no n? `VersionOverrides` do manifesto, inclusive comandos do suplemento, recomendamos usar a implanta??o centralizada por meio do centro de administra??o, se poss?vel. Por padr?o, os comandos do suplemento implantados por meio do cat?logo do SharePoint abrem em um painel de tarefas.

Se voc? est? implantando suplementos em um ambiente local, use um cat?logo do SharePoint. Para saber mais, confira, [Publicar suplementos de conte?do e de painel de tarefas em um cat?logo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Os cat?logos do SharePoint n?o s?o compat?veis com o Office 2016 para Mac. Para implantar Suplementos do Office em clientes do Mac, voc? deve encaminh?-los ao [AppSource]. 

### <a name="outlook-add-in-deployment"></a>Implanta??o de suplemento do Outlook

Em rela??o aos ambientes locais e online que n?o usam o servi?o de identidade do Microsoft Azure AD, ? poss?vel implantar suplementos do Outlook por meio do servidor Exchange. 

Requisitos de implanta??o de suplemento do Outlook:

- Office 365, Exchange Online ou Exchange Server 2013 ou posterior
- Outlook 2013 ou posterior

Para atribuir suplementos a locat?rios, use o Centro de administra??o do Exchange para carregar o manifesto diretamente de um arquivo ou de uma URL ou para adicionar um suplemento do AppSource. Para atribuir suplementos a usu?rios individuais, ? necess?rio usar o Exchange PowerShell. Para saber mais, confira o artigo [Instalar ou remover suplementos do Outlook para a organiza??o](https://technet.microsoft.com/en-us/library/jj943752(v=exchg.150).aspx) no TechNet.

## <a name="see-also"></a>Veja tamb?m

- [Realizar sideload de suplementos do Outlook para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Enviar para o AppSource][AppSource]
- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)
- [Criar listagens eficazes do AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings)
- [Solucionar erros de usu?rios com suplementos do Office](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
