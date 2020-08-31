---
title: Implantar e publicar Suplementos do Office
description: Você pode usar um dos vários métodos para implantar o suplemento do Office para testar ou distribuir aos usuários.
ms.date: 06/02/2020
localization_priority: Priority
ms.openlocfilehash: 7792b2196b686f36b81869ce81a3e3b87659f1c2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292389"
---
# <a name="deploy-and-publish-office-add-ins"></a>Implantar e publicar Suplementos do Office

Você pode usar um dos vários métodos para implantar o suplemento do Office para teste ou distribuição aos usuários.

|**Method**|**Use...**|
|:---------|:------------|
|[Sideload](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)|Usado como parte do processo de desenvolvimento para testar o suplemento em execução no Windows, no iPad, no Mac ou em um navegador. (Não se destina a suplementos de produção.)|
|[Compartilhamento de rede](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)|Como parte do processo de desenvolvimento, teste seu suplemento no Windows após publicá-lo em um servidor que não seja o host local. (Não se destina a suplementos de produção ou para testes no iPad, no Mac ou na Web).|
|[Implantação Centralizada](centralized-deployment.md)|Em uma implantação na nuvem, distribua seu suplemento aos usuários da sua organização usando o Centro de administração do Microsoft 365.|
|[Catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)|Usado para distribuir o suplemento aos usuários da organização em um ambiente local.|
|[AppSource](/office/dev/store/submit-to-appsource-via-partner-center)|Usado para distribuir o suplemento publicamente aos usuários.|
|[Servidor Exchange](#outlook-add-in-deployment)|Usado para distribuir suplementos do Outlook aos usuários em um ambiente local ou online.|

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="deployment-options-by-office-application-and-add-in-type"></a>Opções de implantação por aplicativo do Office e tipo de suplemento

As opções de implantação que estão disponíveis dependem do aplicativo do Office que você pretende usar e o tipo de suplemento que pretende criar.

### <a name="deployment-options-for-word-excel-and-powerpoint-add-ins"></a>Opções de implantação de suplementos para Word, Excel e PowerPoint

| Ponto de extensão | Sideloading | Compartilhamento de rede | Centro de administração do Microsoft 365 |AppSource   | Catálogo do SharePoint\* |
|:----------------|:-----------:|:-------------:|:-----------------------:|:----------:|:--------------------:|
| Conteúdo         | X           | X             | X                       | X          | X                    |
| Painel de tarefas       | X           | X             | X                       | X          | X                    |
| Comando         | X           | X             | X                       | X          |                      |

&#42; Os catálogos do SharePoint não são compatíveis com o Office para Mac.

### <a name="deployment-options-for-outlook-add-ins"></a>Opções de implantação para suplementos do Outlook

| Ponto de extensão | Sideloading | Servidor Exchange | AppSource    |
|:----------------|:-----------:|:---------------:|:------------:|
| Aplicativo de email        | X           | X               | X            |
| Comando         | X           | X               | X            |

## <a name="production-deployment-methods"></a>Métodos de implantação de produção

As seções a seguir fornecem informações adicionais sobre os métodos de implantação mais usados para distribuir os Suplementos do Office de produção aos usuários da organização.

Saiba mais sobre como os usuários finais podem adquirir, inserir e executar suplementos em [Começar a usar seu Suplemento do Office](https://support.office.com/article/start-using-your-office-add-in-82e665c4-6700-4b56-a3f3-ef5441996862).

### <a name="centralized-deployment-via-the-microsoft-365-admin-center"></a>Implantação Centralizada por meio do Centro de administração do Microsoft 365

No Centro de administração do Microsoft 365, é mais fácil para o administrador implantar Suplementos do Office para usuários e grupos da organização. Os suplementos implantados por meio do Centro de administração ficam disponíveis imediatamente para os usuários nos aplicativos do Office, sem a necessidade de configuração do cliente. Você pode usar a Implantação Centralizada para implantar suplementos internos, além de suplementos fornecidos por ISVs.

Para mais informações, confira [Publicar Suplementos do Office usando a Implantação Centralizada por meio do Centro de administração do Microsoft 365](centralized-deployment.md).

### <a name="sharepoint-app-catalog-deployment"></a>Implantação do catálogo de aplicativos do SharePoint

Um catálogo de aplicativos do SharePoint é um conjunto de sites especial que você pode criar para hospedar suplementos do Word, Excel e PowerPoint. Como os catálogos do SharePoint não oferecem suporte a novos recursos de suplementos implementados no nó `VersionOverrides` do manifesto, incluindo comandos de suplementos, recomendamos que você use a Implantação Centralizada por meio do centro de administração, se possível. Comandos de suplemento implantados por meio de um catálogo do SharePoint são abertos em um painel de tarefas por padrão.

Se você está implantando suplementos em um ambiente local, use um catálogo do SharePoint. Para saber mais, confira, [Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

> [!NOTE]
> Os catálogos do SharePoint não são compatíveis com o Office para Mac. Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).

### <a name="outlook-add-in-deployment"></a>Implantação de suplemento do Outlook

Em relação aos ambientes locais e online que não usam o serviço de identidade do Microsoft Azure AD, é possível implantar suplementos do Outlook por meio do servidor Exchange.

Requisitos de implantação de suplemento do Outlook:

- Microsoft 365, Exchange Online ou Exchange Server 2013 ou posterior
- Outlook 2013 ou posterior

Para atribuir suplementos a locatários, use o Centro de administração do Exchange para carregar o manifesto diretamente de um arquivo ou de uma URL ou para adicionar um suplemento do AppSource. Para atribuir suplementos a usuários individuais, é necessário usar o Exchange PowerShell. Para saber mais, confira o artigo [Instalar ou remover suplementos do Outlook para a organização](https://technet.microsoft.com/library/jj943752(v=exchg.150).aspx) no TechNet.

## <a name="see-also"></a>Confira também

- [Realizar sideload de suplementos do Outlook para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Enviar para o AppSource][AppSource]
- [Diretrizes de design para Suplementos do Office](../design/add-in-design.md)
- [Criar listagens eficazes do AppSource](/office/dev/store/create-effective-office-store-listings)
- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)

[AppSource]: /office/dev/store/submit-to-appsource-via-partner-center
