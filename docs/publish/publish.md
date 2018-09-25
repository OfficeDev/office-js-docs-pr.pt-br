---
title: Implantar e publicar seu Suplemento do Office | Documentos Microsoft
description: Métodos e opções para implantar o Suplemento do Office para testes ou distribuição para usuários.
ms.date: 01/23/2018
ms.openlocfilehash: ada786ed7ded1f34d564389c09c2cd5c25c2a331
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004928"
---
# <a name="deploy-and-publish-your-office-add-in"></a>Implantar e publicar seu Suplemento do Office

Você pode usar um dos vários métodos para implantar o suplemento do Office para teste ou distribuição aos usuários.

|**Método**|**Uso...**|
|:---------|:------------|
|[Sideload](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Usado como parte do processo de desenvolvimento para testar o suplemento em execução no Windows, Office Online, iPad ou Mac.|
|[Implantação Centralizada](centralized-deployment.md)|Em uma implantação híbrida ou de nuvem para distribuir seu suplemento aos usuários na sua organização usando o centro de administração do Office 365.|
|[Catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Usado para distribuir o suplemento aos usuários da organização em um ambiente local.|
|[AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)|Usado para distribuir o suplemento publicamente aos usuários.|
|[Servidor Exchange](#outlook-add-in-deployment)|Usado para distribuir suplementos do Outlook aos usuários em um ambiente local ou online.|
|[Compartilhamento de rede](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|No computador do Windows em uma rede na qual você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.|

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)).

## <a name="deployment-options-by-office-host"></a>Opções de implantação pelo host do Office

As opções de implantação disponíveis dependem do host do Office que você pretende usar e do tipo de suplemento que você pretende criar.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Opções de implantação de suplementos para Word, Excel e PowerPoint

| Ponto de extensão | Sideload | Centro de administração do Office 365 |AppSource   | Catálogo do SharePoint\* |
|:----------------|:-----------:|:-----------------------:|:----------:|:--------------------:|
| Conteúdo         | X           | X                       | X          | X                    |
| Painel de tarefas       | X           | X                       | X          | X                    |
| Comando         | X           | X                       | X          |                      |

* Os catálogos do SharePoint não têm suporte para Office para Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Opções de implantação para suplementos do Outlook

| Ponto de extensão | Sideload | Servidor Exchange | AppSource    |
|:----------------|:-----------:|:---------------:|:------------:|
| Aplicativo de email        | X           | X               | X            |
| Comando         | X           | X               | X            |

## <a name="deployment-methods"></a>Métodos de implantação

As seções a seguir fornecem informações adicionais sobre os métodos de implantação mais usados para distribuir Suplementos do Office para usuários em uma organização.

Saiba mais sobre como os usuários finais podem adquirir, inserir e executar suplementos em [Começar a usar seu Suplemento do Office](https://support.office.com/en-ie/article/Start-using-your-Office-Add-in-82e665c4-6700-4b56-a3f3-ef5441996862?ui=en-US&rs=en-IE&ad=IE).

### <a name="centralized-deployment-via-the-office-365-admin-center"></a>Implantação Centralizada por meio do centro de administração do Office 365 

No Centro de administração do Office 365, é mais fácil para o administrador implantar Suplementos do Office para usuários e grupos dentro da organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. É possível usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

Confira mais informações em [Publicar Suplementos do Office usando a Implantação Centralizada por meio do Centro de Administração do Office 365](centralized-deployment.md).

### <a name="sharepoint-catalog-deployment"></a>Implantação de catálogo do SharePoint

O catálogo de suplementos do SharePoint é uma coleção de sites especial que você pode criar para hospedar suplementos dos aplicativos Word, Excel e PowerPoint. Como os catálogos do SharePoint não oferecem suporte para os novos recursos de suplemento implementados no nó `VersionOverrides` do manifesto, inclusive comandos do suplemento, recomendamos usar a implantação centralizada por meio do centro de administração, se possível. Por padrão, os comandos do suplemento implantados por meio do catálogo do SharePoint abrem em um painel de tarefas.

Se você está implantando suplementos em um ambiente local, use um catálogo do SharePoint. Para saber mais, confira, [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Os catálogos do SharePoint não têm suporte para Office para Mac. Para implantar suplementos do Office em clientes Mac, você deve enviá-los para o [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store). 

### <a name="outlook-add-in-deployment"></a>Implantação de suplementos do Outlook

Em relação aos ambientes locais e online que não usam o serviço de identidade do Microsoft Azure AD, é possível implantar suplementos do Outlook por meio do servidor Exchange. 

Requisitos de implantação de suplemento do Outlook:

- Office 365, Exchange Online ou Exchange Server 2013 ou posterior
- Outlook 2013 ou posterior

Para atribuir suplementos a locatários, use o Centro de administração do Exchange para carregar o manifesto diretamente de um arquivo ou de uma URL ou para adicionar um suplemento do AppSource. Para atribuir suplementos a usuários individuais, é necessário usar o Exchange PowerShell. Para saber mais, confira o artigo [Instalar ou remover suplementos do Outlook para a organização](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) no TechNet.

## <a name="see-also"></a>Veja também

- [Realizar sideload de suplementos do Outlook para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Enviar aplicativos para o AppSource][AppSource]
- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)
- [Criar listagens eficazes do AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)
- [Solucionar erros de usuários com suplementos do Office](../testing/testing-and-troubleshooting.md)

[AppSource]: https://docs.microsoft.com/office/dev/store/submit-to-the-office-store
[Office Add-in host and platform availability]: ../overview/office-add-in-availability
