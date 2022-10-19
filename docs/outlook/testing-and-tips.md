---
title: Implante e instale suplementos do Outlook para teste
description: Crie um arquivo de manifesto, implante o arquivo de interface do usuário suplemento em um servidor web, instale o suplemento na caixa de correio e teste o suplemento.
ms.date: 10/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 1b6d29fa85b855adbf75a33345850582d2eecc02
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607517"
---
# <a name="deploy-and-install-outlook-add-ins-for-testing"></a>Implante e instale suplementos do Outlook para teste

Como parte do processo de desenvolvimento de um suplemento do Outlook, você provavelmente se verá implantando e instalando iterativamente o suplemento para teste, o que envolve as etapas a seguir.

1. Criação de um arquivo de manifesto que descreve o suplemento.
1. Implantação dos arquivos da interface do usuário em um servidor Web.
1. Instalação do suplemento em sua caixa de correio.
1. Teste do suplemento, fazendo as alterações apropriadas na interface de usuário ou nos arquivos de manifesto e repetindo as etapas 2 e 3 para testar as alterações.

> [!NOTE]
> [Os painéis personalizados foram preteridos](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/); portanto, certifique-se de estar usando um [ponto de extensão de suplemento com suporte](outlook-add-ins-overview.md#extension-points).

## <a name="create-a-manifest-file-for-the-add-in"></a>Criar um arquivo de manifesto para o suplemento

Cada suplemento é descrito por um manifesto, um documento que fornece ao servidor informações sobre o suplemento, fornece informações descritivas sobre o suplemento para o usuário e identifica o local do arquivo HTML da interface do usuário do suplemento. Você pode armazenar o manifesto em uma pasta ou servidor local, desde que o local possa ser acessado pelo servidor Exchange da caixa de correio que você está testando. Vamos pressupor que você armazena seu manifesto em uma pasta local. Para obter informações sobre como criar um arquivo de manifesto, confira [Manifestos de suplementos do Outlook](manifests.md).

## <a name="deploy-an-add-in-to-a-web-server"></a>Implantar um suplemento em um servidor Web

You can use HTML and JavaScript to create the add-in. The resulting source files are stored on a web server that can be accessed by the Exchange server that hosts the add-in. After initially deploying the source files for the add-in, you can update the add-in UI and behavior by replacing the HTML files or JavaScript files stored on the web server with a new version of the HTML file.

## <a name="install-the-add-in"></a>Instalar o suplemento

Depois de preparar o arquivo de manifesto do suplemento e implantar a interface de usuário do suplemento em um servidor Web que possa ser acessado, é possível realizar o sideload do suplemento para uma caixa de correio em um servidor Exchange usando um cliente do Outlook ou instalar o suplemento executando cmdlets remotos do Windows PowerShell.

### <a name="sideload-the-add-in"></a>Realizar o sideload do suplemento

You can install an add-in if your mailbox is on Exchange Online, Exchange 2013 or a later release. Sideloading add-ins requires at minimum the **My Custom Apps** role for your Exchange Server. In order to test your add-in, or install add-ins in general by specifying a URL or file name for the add-in manifest, you should request your Exchange administrator to provide the necessary permissions.

The Exchange administrator can run the following PowerShell cmdlet to assign a single user the necessary permissions. In this example, `wendyri` is the user's email alias.

```powershell
New-ManagementRoleAssignment -Role "My Custom Apps" -User "wendyri"
```

Se necessário, o administrador poderá executar o seguinte cmdlet para atribuir a vários usuários as permissões semelhantes necessárias.

```powershell
$users = Get-Mailbox *$users | ForEach-Object { New-ManagementRoleAssignment -Role "My Custom Apps" -User $_.Alias}
```

Para saber mais sobre a função Meus Suplementos Personalizados, confira [Função Meus Suplementos Personalizados](/exchange/my-custom-apps-role-exchange-2013-help).

O uso do Microsoft 365 ou do Visual Studio para desenvolver suplementos atribui a você a função de administrador da organização, o que permite que você instale suplementos por arquivo ou URL no EAC, ou por cmdlets do Powershell.

### <a name="install-an-add-in-by-using-remote-powershell"></a>Instalar um suplemento usando o PowerShell remoto

Depois de criar uma sessão remota do Windows PowerShell em seu servidor Exchange, você pode instalar um suplemento do Outlook usando o cmdlet `New-App` com o comando do PowerShell a seguir.

```powershell
New-App -URL:"http://<fully-qualified URL">
```

A URL totalmente qualificada é o local do arquivo de manifesto do suplemento que você preparou para seu suplemento.

Você pode usar os seguintes cmdlets adicionais do Windows PowerShell para gerenciar os suplementos de uma caixa de correio.

- `Get-App` – Lista os suplementos que estão habilitados para uma caixa de correio.
- `Set-App` – Habilita ou desabilita um suplemento em uma caixa de correio.
- `Remove-App` – Remove um suplemento instalado anteriormente de um servidor Exchange.

## <a name="client-versions"></a>Versões de cliente

A decisão de quais versões de cliente do Outlook testar depende dos seus requisitos de desenvolvimento.

- If you're developing an add-in for private use, or only for members of your organization, then it is important to test the versions of Outlook that your company uses. Keep in mind that some users may use Outlook on the web, so testing your company's standard browser versions is also important.

- If you're developing an add-in to list in [AppSource](https://appsource.microsoft.com), you must test the required versions as specified in the [Commercial marketplace certification policies 1120.3](/legal/marketplace/certification-policies#11203-functionality). This includes:
  - A versão mais recente do Outlook no Windows e a versão anterior à mais recente.
  - A versão mais recente do Outlook no Mac.
  - A versão mais recente do Outlook no iOS e Android (se o suplemento for [compatível com mobilidade](add-mobile-support.md)).
  - As versões do navegador especificadas na política de validação do mercado comercial 1120.3

> [!NOTE]
> Se seu suplemento não for compatível com um dos clientes acima devido a uma [solicitação de um conjunto de exigências da API](apis.md) que o cliente não dá suporte, esse cliente será removido da lista de clientes exigidos.

## <a name="outlook-on-the-web-and-exchange-server-versions"></a>Versões do Outlook na Web e do Exchange Server

Os usuários de contas do Microsoft 365 e de consumidor veem a versão moderna da interface do usuário ao acessar o Outlook na Web e não veem mais a versão clássica que foi substituída. No entanto, os servidores locais do Exchange continuam oferecendo suporte ao Outlook na Web clássico. Portanto, durante o processo de validação, seu envio poderá receber um aviso de que o suplemento não é compatível com o Outlook na Web clássico. Nesse caso, considere testar o suplemento em um ambiente do Exchange local. Esse aviso não bloqueará seu envio ao AppSource, mas seus clientes poderão ter uma experiência abaixo do ideal, caso usem o Outlook na Web em um ambiente do Exchange local.

Para atenuar isso, é recomendável que se faça o teste do suplemento no Outlook na Web conectado ao seu próprio ambiente Exchange local. Para saber mais, confira as orientações sobre como [Estabelecer um ambiente de teste do Exchange 2016 ou do Exchange 2019](/Exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019&preserve-view=true#establish-an-exchange-2016-or-exchange-2019-test-environment) e como gerenciar o [Outlook na Web no Exchange Server](/exchange/clients/outlook-on-the-web/outlook-on-the-web?view=exchserver-2019&preserve-view=true).

Como alternativa, você pode optar por pagar e usar um serviço que hospeda e gerencia servidores locais do Exchange. Algumas das opções são:

- [Rackspace](https://www.rackspace.com/email-hosting/exchange-server)
- [Hostway](https://hostway.com/microsoft-exchange/)

Além disso, se você não deseja que seus suplementos estejam disponíveis para usuários conectados ao Exchange local, é possível definir o [conjunto de requisitos](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#exchange-server-support) no manifesto de suplemento como 1.6 ou superior. Esses suplementos não serão testados nem validados na interface do usuário do Outlook na Web clássico.

## <a name="see-also"></a>Confira também

- [Solucionar erros de usuários com Suplementos do Office](../testing/testing-and-troubleshooting.md)
